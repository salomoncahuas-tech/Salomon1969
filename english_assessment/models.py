"""
Modelos de base de datos para la aplicación de Evaluaciones de Inglés.
Soporta grados (1ro - 6to Primaria, 1ro - 5to Secundaria), niveles (Beginner, Elementary,
Pre-Intermediate, Intermediate, Upper-Intermediate), evaluaciones, preguntas y resultados.
"""

from datetime import datetime
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()


class Teacher(UserMixin, db.Model):
    """Modelo de profesor/administrador."""
    __tablename__ = 'teachers'

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    full_name = db.Column(db.String(150), nullable=False)
    email = db.Column(db.String(150))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    assessments = db.relationship('Assessment', backref='teacher', lazy=True)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)


class Grade(db.Model):
    """Grado escolar (1ro Primaria, 2do Secundaria, etc.)."""
    __tablename__ = 'grades'

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), nullable=False)
    level_type = db.Column(db.String(20), nullable=False)  # 'primaria' o 'secundaria'
    order = db.Column(db.Integer, nullable=False)

    assessments = db.relationship('Assessment', backref='grade', lazy=True)

    def __repr__(self):
        return f'{self.name}'


class EnglishLevel(db.Model):
    """Nivel de inglés (Beginner, Elementary, etc.)."""
    __tablename__ = 'english_levels'

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), nullable=False)
    code = db.Column(db.String(10), nullable=False)
    description = db.Column(db.String(200))
    order = db.Column(db.Integer, nullable=False)

    assessments = db.relationship('Assessment', backref='english_level', lazy=True)

    def __repr__(self):
        return f'{self.name} ({self.code})'


class Assessment(db.Model):
    """Evaluación o tarea de inglés."""
    __tablename__ = 'assessments'

    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text)
    assessment_type = db.Column(db.String(20), nullable=False)  # 'exam', 'homework', 'quiz', 'practice'
    grade_id = db.Column(db.Integer, db.ForeignKey('grades.id'), nullable=False)
    english_level_id = db.Column(db.Integer, db.ForeignKey('english_levels.id'), nullable=False)
    teacher_id = db.Column(db.Integer, db.ForeignKey('teachers.id'), nullable=False)
    time_limit_minutes = db.Column(db.Integer, default=0)  # 0 = sin límite
    is_active = db.Column(db.Boolean, default=True)
    shuffle_questions = db.Column(db.Boolean, default=False)
    show_results = db.Column(db.Boolean, default=True)
    access_code = db.Column(db.String(20), unique=True, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    due_date = db.Column(db.DateTime)

    questions = db.relationship('Question', backref='assessment', lazy=True,
                                order_by='Question.order', cascade='all, delete-orphan')
    results = db.relationship('StudentResult', backref='assessment', lazy=True,
                              cascade='all, delete-orphan')

    @property
    def total_points(self):
        return sum(q.points for q in self.questions)

    @property
    def question_count(self):
        return len(self.questions)


class Question(db.Model):
    """Pregunta dentro de una evaluación."""
    __tablename__ = 'questions'

    id = db.Column(db.Integer, primary_key=True)
    assessment_id = db.Column(db.Integer, db.ForeignKey('assessments.id'), nullable=False)
    question_type = db.Column(db.String(30), nullable=False)
    # Tipos: 'multiple_choice', 'true_false', 'fill_blank', 'matching', 'ordering', 'short_answer'
    text = db.Column(db.Text, nullable=False)
    instruction = db.Column(db.Text)  # Instrucción específica para la pregunta
    points = db.Column(db.Float, default=1.0)
    order = db.Column(db.Integer, default=0)
    media_url = db.Column(db.String(500))  # URL de imagen o audio opcional

    options = db.relationship('QuestionOption', backref='question', lazy=True,
                              order_by='QuestionOption.order', cascade='all, delete-orphan')
    answers = db.relationship('StudentAnswer', backref='question', lazy=True,
                              cascade='all, delete-orphan')


class QuestionOption(db.Model):
    """Opción de respuesta para preguntas de opción múltiple, matching, etc."""
    __tablename__ = 'question_options'

    id = db.Column(db.Integer, primary_key=True)
    question_id = db.Column(db.Integer, db.ForeignKey('questions.id'), nullable=False)
    text = db.Column(db.Text, nullable=False)
    is_correct = db.Column(db.Boolean, default=False)
    match_text = db.Column(db.Text)  # Para preguntas tipo matching
    order = db.Column(db.Integer, default=0)


class StudentResult(db.Model):
    """Resultado global de un alumno en una evaluación."""
    __tablename__ = 'student_results'

    id = db.Column(db.Integer, primary_key=True)
    assessment_id = db.Column(db.Integer, db.ForeignKey('assessments.id'), nullable=False)
    student_name = db.Column(db.String(150), nullable=False)
    student_code = db.Column(db.String(50))  # Código de alumno opcional
    grade_section = db.Column(db.String(20))  # Sección (A, B, C...)
    score = db.Column(db.Float, default=0)
    total_points = db.Column(db.Float, default=0)
    percentage = db.Column(db.Float, default=0)
    started_at = db.Column(db.DateTime, default=datetime.utcnow)
    completed_at = db.Column(db.DateTime)
    is_completed = db.Column(db.Boolean, default=False)

    answers = db.relationship('StudentAnswer', backref='result', lazy=True,
                              cascade='all, delete-orphan')


class StudentAnswer(db.Model):
    """Respuesta individual de un alumno a una pregunta."""
    __tablename__ = 'student_answers'

    id = db.Column(db.Integer, primary_key=True)
    result_id = db.Column(db.Integer, db.ForeignKey('student_results.id'), nullable=False)
    question_id = db.Column(db.Integer, db.ForeignKey('questions.id'), nullable=False)
    answer_text = db.Column(db.Text)
    selected_option_id = db.Column(db.Integer, db.ForeignKey('question_options.id'))
    is_correct = db.Column(db.Boolean, default=False)
    points_earned = db.Column(db.Float, default=0)

    selected_option = db.relationship('QuestionOption', foreign_keys=[selected_option_id])
