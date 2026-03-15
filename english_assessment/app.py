"""
English Assessment Web Application
Aplicación web para evaluaciones y tareas de inglés.
Permite a alumnos de distintos grados y niveles acceder a evaluaciones mediante un enlace.
"""

import os
import logging
import secrets
import random
from datetime import datetime
from functools import wraps

from flask import (Flask, render_template, request, redirect, url_for,
                   flash, session, jsonify, abort, send_from_directory)
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from werkzeug.utils import secure_filename

from models import (db, Teacher, Grade, EnglishLevel, Assessment, Question,
                    QuestionOption, StudentResult, StudentAnswer)

basedir = os.path.abspath(os.path.dirname(__file__))

# Audio upload configuration
AUDIO_UPLOAD_FOLDER = os.path.join(basedir, 'static', 'audio')
ALLOWED_AUDIO_EXTENSIONS = {'mp3', 'wav', 'ogg', 'webm', 'm4a'}
MAX_AUDIO_SIZE_MB = 15  # Max 15 MB per audio file (enough for 3 min)


def allowed_audio_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_AUDIO_EXTENSIONS


def get_or_create_secret_key():
    """Read or generate a stable SECRET_KEY so sessions survive restarts."""
    key_file = os.path.join(basedir, '.secret_key')
    if os.path.exists(key_file):
        with open(key_file, 'r') as f:
            return f.read().strip()
    key = secrets.token_hex(32)
    with open(key_file, 'w') as f:
        f.write(key)
    return key


app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY') or get_or_create_secret_key()

# Database configuration: use PostgreSQL on Render, SQLite locally
database_url = os.environ.get('DATABASE_URL', f'sqlite:///{os.path.join(basedir, "english_assessments.db")}')
# Render provides postgres:// but SQLAlchemy requires postgresql://
if database_url.startswith('postgres://'):
    database_url = database_url.replace('postgres://', 'postgresql://', 1)
app.config['SQLALCHEMY_DATABASE_URI'] = database_url
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['MAX_CONTENT_LENGTH'] = MAX_AUDIO_SIZE_MB * 1024 * 1024

# Ensure audio upload directory exists
os.makedirs(AUDIO_UPLOAD_FOLDER, exist_ok=True)

# Configure logging for production
if not app.debug:
    logging.basicConfig(level=logging.INFO)
    app.logger.setLevel(logging.INFO)

db.init_app(app)
login_manager = LoginManager(app)
login_manager.login_view = 'admin_login'
login_manager.login_message = 'Please log in to access the admin panel.'


@login_manager.user_loader
def load_user(user_id):
    return db.session.get(Teacher, int(user_id))


# ---------------------------------------------------------------------------
# PUBLIC ROUTES - Student access
# ---------------------------------------------------------------------------

@app.route('/')
def index():
    """Landing page with access code form only."""
    return render_template('index.html')


@app.route('/assessments')
def list_assessments():
    """Redirect to home - assessments are only accessible via access code."""
    return redirect(url_for('index'))


@app.route('/access', methods=['GET', 'POST'])
def access_by_code():
    """Access an assessment by its unique code."""
    if request.method == 'POST':
        code = request.form.get('access_code', '').strip().upper()
        assessment = Assessment.query.filter_by(access_code=code, is_active=True).first()
        if assessment:
            session[f'access_{assessment.id}'] = True
            return redirect(url_for('start_assessment', assessment_id=assessment.id))
        flash('Invalid access code. Please verify and try again.', 'error')
    return render_template('access_code.html')


@app.route('/assessment/<int:assessment_id>')
def start_assessment(assessment_id):
    """Assessment info page before starting - requires access code."""
    assessment = Assessment.query.get_or_404(assessment_id)
    if not assessment.is_active:
        abort(404)
    if not session.get(f'access_{assessment_id}'):
        flash('Please enter the access code to start this assessment.', 'error')
        return redirect(url_for('access_by_code'))
    return render_template('start_assessment.html', assessment=assessment)


@app.route('/assessment/<int:assessment_id>/take', methods=['GET', 'POST'])
def take_assessment(assessment_id):
    """Take an assessment - display questions and collect answers."""
    assessment = Assessment.query.get_or_404(assessment_id)
    if not assessment.is_active:
        abort(404)
    if not session.get(f'access_{assessment_id}'):
        flash('Please enter the access code to start this assessment.', 'error')
        return redirect(url_for('access_by_code'))

    if request.method == 'POST':
        student_name = request.form.get('student_name', '').strip()
        student_code = request.form.get('student_code', '').strip()
        grade_section = request.form.get('grade_section', '').strip()

        if not student_name:
            flash('Please enter your name.', 'error')
            return redirect(url_for('start_assessment', assessment_id=assessment_id))

        # Create student result
        try:
            result = StudentResult(
                assessment_id=assessment_id,
                student_name=student_name,
                student_code=student_code,
                grade_section=grade_section,
                total_points=assessment.total_points,
                started_at=datetime.utcnow()
            )
            db.session.add(result)
            db.session.commit()
        except Exception as e:
            db.session.rollback()
            app.logger.error(
                'Error creating student result for assessment %d: %s',
                assessment_id, str(e)
            )
            flash('There was an error starting the assessment. Please try again.', 'error')
            return redirect(url_for('start_assessment', assessment_id=assessment_id))

        session[f'result_{assessment_id}'] = result.id

        questions = list(assessment.questions)
        if assessment.shuffle_questions:
            random.shuffle(questions)

        return render_template('take_assessment.html',
                               assessment=assessment,
                               questions=questions,
                               result=result)

    return redirect(url_for('start_assessment', assessment_id=assessment_id))


@app.route('/assessment/<int:assessment_id>/submit', methods=['POST'])
def submit_assessment(assessment_id):
    """Process submitted answers and calculate score."""
    assessment = Assessment.query.get_or_404(assessment_id)
    result_id = session.get(f'result_{assessment_id}')
    if not result_id:
        # Fallback: recover result_id from the hidden form field
        result_id = request.form.get('result_id', type=int)

    if not result_id:
        flash('Session expired. Please start the assessment again.', 'error')
        return redirect(url_for('start_assessment', assessment_id=assessment_id))

    result = db.session.get(StudentResult, result_id)
    if not result or result.assessment_id != assessment_id:
        flash('Invalid submission. Please start the assessment again.', 'error')
        return redirect(url_for('start_assessment', assessment_id=assessment_id))
    if result.is_completed:
        flash('This assessment has already been submitted.', 'error')
        return redirect(url_for('start_assessment', assessment_id=assessment_id))

    try:
        total_score = 0

        for question in assessment.questions:
            answer_key = f'question_{question.id}'

            if question.question_type in ('multiple_choice', 'listening_multiple_choice'):
                selected_id = request.form.get(answer_key, type=int)
                selected_option = db.session.get(QuestionOption, selected_id) if selected_id else None
                is_correct = selected_option.is_correct if selected_option else False
                points = question.points if is_correct else 0
                total_score += points

                student_answer = StudentAnswer(
                    result_id=result.id,
                    question_id=question.id,
                    selected_option_id=selected_id,
                    answer_text=selected_option.text if selected_option else '',
                    is_correct=is_correct,
                    points_earned=points
                )

            elif question.question_type in ('true_false', 'listening_true_false'):
                answer_val = request.form.get(answer_key, '').strip().lower()
                correct_option = next((o for o in question.options if o.is_correct), None)
                correct_val = correct_option.text.strip().lower() if correct_option else ''
                is_correct = answer_val == correct_val
                points = question.points if is_correct else 0
                total_score += points

                student_answer = StudentAnswer(
                    result_id=result.id,
                    question_id=question.id,
                    answer_text=answer_val,
                    is_correct=is_correct,
                    points_earned=points
                )

            elif question.question_type in ('fill_blank', 'listening_fill_blank'):
                answer_text = request.form.get(answer_key, '').strip()
                correct_option = next((o for o in question.options if o.is_correct), None)
                correct_text = correct_option.text.strip() if correct_option else ''
                is_correct = answer_text.lower() == correct_text.lower()
                points = question.points if is_correct else 0
                total_score += points

                student_answer = StudentAnswer(
                    result_id=result.id,
                    question_id=question.id,
                    answer_text=answer_text,
                    is_correct=is_correct,
                    points_earned=points
                )

            elif question.question_type == 'matching':
                correct_count = 0
                total_pairs = len([o for o in question.options if o.match_text])
                answers_parts = []
                for option in question.options:
                    if option.match_text:
                        match_key = f'match_{question.id}_{option.id}'
                        student_match = request.form.get(match_key, '').strip()
                        answers_parts.append(f'{option.id}:{student_match}')
                        if student_match == option.match_text:
                            correct_count += 1
                points = (correct_count / total_pairs * question.points) if total_pairs > 0 else 0
                total_score += points

                student_answer = StudentAnswer(
                    result_id=result.id,
                    question_id=question.id,
                    answer_text='|'.join(answers_parts),
                    is_correct=correct_count == total_pairs,
                    points_earned=points
                )

            elif question.question_type == 'ordering':
                answer_text = request.form.get(answer_key, '').strip()
                correct_order = ','.join(str(o.id) for o in sorted(question.options, key=lambda x: x.order))
                is_correct = answer_text == correct_order
                points = question.points if is_correct else 0
                total_score += points

                student_answer = StudentAnswer(
                    result_id=result.id,
                    question_id=question.id,
                    answer_text=answer_text,
                    is_correct=is_correct,
                    points_earned=points
                )

            else:  # short_answer, listening_short_answer
                answer_text = request.form.get(answer_key, '').strip()
                student_answer = StudentAnswer(
                    result_id=result.id,
                    question_id=question.id,
                    answer_text=answer_text,
                    is_correct=False,
                    points_earned=0
                )

            db.session.add(student_answer)

        result.score = total_score
        result.percentage = (total_score / result.total_points * 100) if result.total_points > 0 else 0
        result.is_completed = True
        result.completed_at = datetime.utcnow()
        db.session.commit()

        session.pop(f'result_{assessment_id}', None)

        return redirect(url_for('show_results', result_id=result.id))

    except Exception as e:
        db.session.rollback()
        app.logger.error(
            'Error submitting assessment %d for result %d: %s',
            assessment_id, result_id, str(e)
        )
        flash(
            'There was an error submitting your assessment. '
            'Please try again. If the problem persists, contact your teacher.',
            'error'
        )
        return redirect(url_for('start_assessment', assessment_id=assessment_id))


@app.route('/results/<int:result_id>')
def show_results(result_id):
    """Show assessment results to the student."""
    result = StudentResult.query.get_or_404(result_id)
    if not result.assessment.show_results:
        return render_template('results_hidden.html', result=result)
    return render_template('results.html', result=result)


# ---------------------------------------------------------------------------
# ADMIN ROUTES - Teacher panel
# ---------------------------------------------------------------------------

@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    if current_user.is_authenticated:
        return redirect(url_for('admin_dashboard'))
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        teacher = Teacher.query.filter_by(username=username).first()
        if teacher and teacher.check_password(password):
            login_user(teacher)
            return redirect(url_for('admin_dashboard'))
        flash('Invalid username or password.', 'error')
    return render_template('admin/login.html')


@app.route('/admin/logout')
@login_required
def admin_logout():
    logout_user()
    return redirect(url_for('index'))


@app.route('/admin')
@login_required
def admin_dashboard():
    assessments = Assessment.query.filter_by(teacher_id=current_user.id)\
        .order_by(Assessment.created_at.desc()).all()
    total_results = StudentResult.query.join(Assessment)\
        .filter(Assessment.teacher_id == current_user.id, StudentResult.is_completed == True).count()
    return render_template('admin/dashboard.html',
                           assessments=assessments,
                           total_results=total_results)


@app.route('/admin/assessment/new', methods=['GET', 'POST'])
@login_required
def admin_new_assessment():
    grades = Grade.query.order_by(Grade.order).all()
    levels = EnglishLevel.query.order_by(EnglishLevel.order).all()

    if request.method == 'POST':
        access_code = secrets.token_hex(4).upper()
        while Assessment.query.filter_by(access_code=access_code).first():
            access_code = secrets.token_hex(4).upper()

        due_date_str = request.form.get('due_date', '').strip()
        due_date = datetime.strptime(due_date_str, '%Y-%m-%d') if due_date_str else None

        assessment = Assessment(
            title=request.form['title'],
            description=request.form.get('description', ''),
            assessment_type=request.form['assessment_type'],
            grade_id=int(request.form['grade_id']),
            english_level_id=int(request.form['english_level_id']),
            teacher_id=current_user.id,
            time_limit_minutes=int(request.form.get('time_limit', 0)),
            shuffle_questions='shuffle' in request.form,
            show_results='show_results' in request.form,
            access_code=access_code,
            due_date=due_date
        )
        db.session.add(assessment)
        db.session.commit()
        flash(f'Assessment created! Access code: {access_code}', 'success')
        return redirect(url_for('admin_edit_questions', assessment_id=assessment.id))

    return render_template('admin/new_assessment.html', grades=grades, levels=levels)


@app.route('/admin/assessment/<int:assessment_id>/questions', methods=['GET', 'POST'])
@login_required
def admin_edit_questions(assessment_id):
    assessment = Assessment.query.get_or_404(assessment_id)
    if assessment.teacher_id != current_user.id:
        abort(403)

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'add_question':
            q_type = request.form['question_type']
            text = request.form['question_text']
            instruction = request.form.get('instruction', '')
            points = float(request.form.get('points', 1))
            order = len(assessment.questions) + 1

            # Handle audio URL or file upload
            media_url = None
            media_type = 'image'
            audio_transcript = None

            if q_type.startswith('listening_'):
                media_type = 'audio'
                audio_transcript = request.form.get('audio_transcript', '').strip() or None

                # Check for audio URL first
                audio_url = request.form.get('audio_url', '').strip()
                if audio_url:
                    media_url = audio_url
                else:
                    # Check for uploaded audio file
                    audio_file = request.files.get('audio_file')
                    if audio_file and audio_file.filename and allowed_audio_file(audio_file.filename):
                        filename = secure_filename(audio_file.filename)
                        # Add timestamp to avoid collisions
                        name, ext = os.path.splitext(filename)
                        filename = f'{name}_{secrets.token_hex(4)}{ext}'
                        audio_file.save(os.path.join(AUDIO_UPLOAD_FOLDER, filename))
                        media_url = url_for('static', filename=f'audio/{filename}')

            question = Question(
                assessment_id=assessment_id,
                question_type=q_type,
                text=text,
                instruction=instruction,
                points=points,
                order=order,
                media_url=media_url,
                media_type=media_type,
                audio_transcript=audio_transcript
            )
            db.session.add(question)
            db.session.commit()

            # Add options
            if q_type in ('multiple_choice', 'true_false', 'fill_blank', 'matching', 'ordering',
                          'listening_multiple_choice', 'listening_true_false', 'listening_fill_blank'):
                option_texts = request.form.getlist('option_text[]')
                correct_indices = request.form.getlist('correct[]')
                match_texts = request.form.getlist('match_text[]')

                for i, opt_text in enumerate(option_texts):
                    if not opt_text.strip():
                        continue
                    option = QuestionOption(
                        question_id=question.id,
                        text=opt_text.strip(),
                        is_correct=(str(i) in correct_indices),
                        match_text=match_texts[i].strip() if i < len(match_texts) and match_texts[i].strip() else None,
                        order=i
                    )
                    db.session.add(option)
                db.session.commit()

            flash('Question added successfully.', 'success')

        elif action == 'delete_question':
            q_id = int(request.form['question_id'])
            question = Question.query.get_or_404(q_id)
            if question.assessment_id == assessment_id:
                db.session.delete(question)
                db.session.commit()
                flash('Question deleted.', 'success')

        return redirect(url_for('admin_edit_questions', assessment_id=assessment_id))

    return render_template('admin/edit_questions.html', assessment=assessment)


@app.route('/admin/assessment/<int:assessment_id>/toggle', methods=['POST'])
@login_required
def admin_toggle_assessment(assessment_id):
    assessment = Assessment.query.get_or_404(assessment_id)
    if assessment.teacher_id != current_user.id:
        abort(403)
    assessment.is_active = not assessment.is_active
    db.session.commit()
    status = 'activated' if assessment.is_active else 'deactivated'
    flash(f'Assessment {status}.', 'success')
    return redirect(url_for('admin_dashboard'))


@app.route('/admin/assessment/<int:assessment_id>/delete', methods=['POST'])
@login_required
def admin_delete_assessment(assessment_id):
    assessment = Assessment.query.get_or_404(assessment_id)
    if assessment.teacher_id != current_user.id:
        abort(403)
    db.session.delete(assessment)
    db.session.commit()
    flash('Assessment deleted.', 'success')
    return redirect(url_for('admin_dashboard'))


@app.route('/admin/assessment/<int:assessment_id>/results')
@login_required
def admin_view_results(assessment_id):
    assessment = Assessment.query.get_or_404(assessment_id)
    if assessment.teacher_id != current_user.id:
        abort(403)
    results = StudentResult.query.filter_by(
        assessment_id=assessment_id, is_completed=True
    ).order_by(StudentResult.completed_at.desc()).all()
    return render_template('admin/view_results.html',
                           assessment=assessment, results=results)


@app.route('/admin/result/<int:result_id>/detail')
@login_required
def admin_result_detail(result_id):
    result = StudentResult.query.get_or_404(result_id)
    if result.assessment.teacher_id != current_user.id:
        abort(403)
    return render_template('admin/result_detail.html', result=result)


# ---------------------------------------------------------------------------
# API endpoints for dynamic UI
# ---------------------------------------------------------------------------

@app.route('/api/assessments')
@login_required
def api_assessments():
    query = Assessment.query.filter_by(teacher_id=current_user.id)
    assessments = query.order_by(Assessment.created_at.desc()).all()
    return jsonify([{
        'id': a.id,
        'title': a.title,
        'type': a.assessment_type,
        'grade': str(a.grade),
        'level': str(a.english_level),
        'questions': a.question_count,
        'points': a.total_points,
        'code': a.access_code,
        'due_date': a.due_date.strftime('%Y-%m-%d') if a.due_date else None
    } for a in assessments])


# ---------------------------------------------------------------------------
# Error handlers
# ---------------------------------------------------------------------------

@app.errorhandler(404)
def not_found(e):
    return render_template('error.html', code=404,
                           message='Page not found.'), 404


@app.errorhandler(403)
def forbidden(e):
    return render_template('error.html', code=403,
                           message='Access denied.'), 403


@app.errorhandler(500)
def internal_error(e):
    db.session.rollback()
    app.logger.error('Internal server error: %s', str(e))
    return render_template('error.html', code=500,
                           message='An unexpected error occurred. Please try again later.'), 500


# ---------------------------------------------------------------------------
# App initialization
# ---------------------------------------------------------------------------

def init_db():
    """Initialize database with default data."""
    db.create_all()

    if Grade.query.first():
        return

    # Grados de Primaria
    for i, name in enumerate(['1st Grade Primary', '2nd Grade Primary', '3rd Grade Primary',
                               '4th Grade Primary', '5th Grade Primary', '6th Grade Primary'], 1):
        db.session.add(Grade(name=name, level_type='primaria', order=i))

    # Grados de Secundaria
    for i, name in enumerate(['1st Grade Secondary', '2nd Grade Secondary', '3rd Grade Secondary',
                               '4th Grade Secondary', '5th Grade Secondary'], 7):
        db.session.add(Grade(name=name, level_type='secundaria', order=i))

    # Niveles de inglés
    levels_data = [
        ('Beginner', 'A1', 'Basic understanding of simple phrases and expressions.'),
        ('Elementary', 'A2', 'Can communicate in simple and routine tasks.'),
        ('Pre-Intermediate', 'B1', 'Can deal with most situations while traveling.'),
        ('Intermediate', 'B2', 'Can interact with a degree of fluency and spontaneity.'),
        ('Upper-Intermediate', 'C1', 'Can express ideas fluently and spontaneously.'),
    ]
    for i, (name, code, desc) in enumerate(levels_data, 1):
        db.session.add(EnglishLevel(name=name, code=code, description=desc, order=i))

    # Default admin teacher
    admin = Teacher(username='admin', full_name='Administrator', email='admin@school.edu')
    admin.set_password('admin123')
    db.session.add(admin)

    db.session.commit()


with app.app_context():
    init_db()
    if not Assessment.query.first():
        try:
            from seed_data import seed_assessments
            seed_assessments()
        except Exception as e:
            print(f'Warning: Could not auto-seed assessments: {e}')
    # Add new assessments that may not exist in older databases
    if not Assessment.query.filter_by(access_code='FREQ2SEC').first():
        try:
            from seed_data import seed_freq2sec
            seed_freq2sec()
            print('Added new practice: Adverbs of Frequency (FREQ2SEC)')
        except Exception as e:
            print(f'Warning: Could not add FREQ2SEC: {e}')
    # Add listening comprehension assessments
    if not Assessment.query.filter_by(access_code='LIST3PA1').first():
        try:
            from seed_data import seed_listening_comprehension
            seed_listening_comprehension()
            print('Added new: Listening Comprehension assessments')
        except Exception as e:
            print(f'Warning: Could not add Listening Comprehension: {e}')


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
