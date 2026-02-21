"""
Seed data script for English Assessment Platform.
Populates the database with sample assessments for various grades and levels.
Run: python seed_data.py
"""

import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

from app import app, db
from models import (Teacher, Grade, EnglishLevel, Assessment, Question,
                    QuestionOption)


def seed_assessments():
    """Create sample assessments with questions for different grades and levels."""
    with app.app_context():
        # Check if we already have assessments
        if Assessment.query.first():
            print('Assessments already exist. Skipping seed.')
            return

        teacher = Teacher.query.filter_by(username='admin').first()
        if not teacher:
            print('No admin teacher found. Run the app first to initialize the database.')
            return

        # ================================================================
        # ASSESSMENT 1: 3rd Grade Primary - Beginner (A1) - Vocabulary Quiz
        # ================================================================
        grade_3p = Grade.query.filter_by(order=3).first()
        level_a1 = EnglishLevel.query.filter_by(code='A1').first()

        a1 = Assessment(
            title='Colors and Animals Vocabulary Quiz',
            description='Test your knowledge of basic colors and animal names in English.',
            assessment_type='quiz',
            grade_id=grade_3p.id,
            english_level_id=level_a1.id,
            teacher_id=teacher.id,
            time_limit_minutes=15,
            show_results=True,
            access_code='COLORS01'
        )
        db.session.add(a1)
        db.session.commit()

        # Q1 - Multiple choice
        q = Question(assessment_id=a1.id, question_type='multiple_choice',
                     text='What color is the sky on a clear day?',
                     instruction='Choose the correct answer.', points=2, order=1)
        db.session.add(q)
        db.session.commit()
        for text, correct in [('Red', False), ('Blue', True), ('Green', False), ('Yellow', False)]:
            db.session.add(QuestionOption(question_id=q.id, text=text, is_correct=correct, order=0))

        # Q2 - Multiple choice
        q = Question(assessment_id=a1.id, question_type='multiple_choice',
                     text='Which animal says "meow"?',
                     instruction='Choose the correct answer.', points=2, order=2)
        db.session.add(q)
        db.session.commit()
        for text, correct in [('Dog', False), ('Cat', True), ('Bird', False), ('Fish', False)]:
            db.session.add(QuestionOption(question_id=q.id, text=text, is_correct=correct, order=0))

        # Q3 - True/False
        q = Question(assessment_id=a1.id, question_type='true_false',
                     text='A dog is a pet animal.',
                     instruction='Select True or False.', points=1, order=3)
        db.session.add(q)
        db.session.commit()
        db.session.add(QuestionOption(question_id=q.id, text='true', is_correct=True, order=0))
        db.session.add(QuestionOption(question_id=q.id, text='false', is_correct=False, order=1))

        # Q4 - Fill in the blank
        q = Question(assessment_id=a1.id, question_type='fill_blank',
                     text='The color of grass is ______.',
                     instruction='Type the missing word.', points=2, order=4)
        db.session.add(q)
        db.session.commit()
        db.session.add(QuestionOption(question_id=q.id, text='green', is_correct=True, order=0))

        # Q5 - Matching
        q = Question(assessment_id=a1.id, question_type='matching',
                     text='Match each animal with its sound.',
                     instruction='Match the items on the left with the correct option on the right.',
                     points=3, order=5)
        db.session.add(q)
        db.session.commit()
        matches = [('Dog', 'Woof'), ('Cat', 'Meow'), ('Cow', 'Moo')]
        for i, (item, match) in enumerate(matches):
            db.session.add(QuestionOption(question_id=q.id, text=item, match_text=match,
                                          is_correct=True, order=i))

        # Q6 - Multiple choice
        q = Question(assessment_id=a1.id, question_type='multiple_choice',
                     text='How many legs does a spider have?',
                     instruction='Choose the correct answer.', points=2, order=6)
        db.session.add(q)
        db.session.commit()
        for text, correct in [('4', False), ('6', False), ('8', True), ('10', False)]:
            db.session.add(QuestionOption(question_id=q.id, text=text, is_correct=correct, order=0))

        db.session.commit()
        print(f'Created: {a1.title} (Code: {a1.access_code})')

        # ================================================================
        # ASSESSMENT 2: 1st Grade Secondary - Elementary (A2) - Grammar
        # ================================================================
        grade_1s = Grade.query.filter_by(order=7).first()
        level_a2 = EnglishLevel.query.filter_by(code='A2').first()

        a2 = Assessment(
            title='Present Simple and Continuous Grammar Test',
            description='Evaluate your understanding of present simple and present continuous tenses.',
            assessment_type='exam',
            grade_id=grade_1s.id,
            english_level_id=level_a2.id,
            teacher_id=teacher.id,
            time_limit_minutes=30,
            show_results=True,
            access_code='GRAMMAR01'
        )
        db.session.add(a2)
        db.session.commit()

        # Q1
        q = Question(assessment_id=a2.id, question_type='multiple_choice',
                     text='She _____ to school every day.',
                     instruction='Choose the correct form of the verb.', points=2, order=1)
        db.session.add(q)
        db.session.commit()
        for text, correct in [('go', False), ('goes', True), ('going', False), ('is go', False)]:
            db.session.add(QuestionOption(question_id=q.id, text=text, is_correct=correct, order=0))

        # Q2
        q = Question(assessment_id=a2.id, question_type='multiple_choice',
                     text='They _____ football right now.',
                     instruction='Choose the correct form of the verb.', points=2, order=2)
        db.session.add(q)
        db.session.commit()
        for text, correct in [('play', False), ('plays', False), ('are playing', True), ('is playing', False)]:
            db.session.add(QuestionOption(question_id=q.id, text=text, is_correct=correct, order=0))

        # Q3
        q = Question(assessment_id=a2.id, question_type='true_false',
                     text='We use present continuous for actions happening right now.',
                     instruction='Select True or False.', points=1, order=3)
        db.session.add(q)
        db.session.commit()
        db.session.add(QuestionOption(question_id=q.id, text='true', is_correct=True, order=0))
        db.session.add(QuestionOption(question_id=q.id, text='false', is_correct=False, order=1))

        # Q4
        q = Question(assessment_id=a2.id, question_type='fill_blank',
                     text='He _____ (not/like) coffee. (Use present simple negative)',
                     instruction='Type the correct form.', points=2, order=4)
        db.session.add(q)
        db.session.commit()
        db.session.add(QuestionOption(question_id=q.id, text="doesn't like", is_correct=True, order=0))

        # Q5
        q = Question(assessment_id=a2.id, question_type='multiple_choice',
                     text='_____ you usually wake up early?',
                     instruction='Choose the correct auxiliary verb.', points=2, order=5)
        db.session.add(q)
        db.session.commit()
        for text, correct in [('Do', True), ('Does', False), ('Are', False), ('Is', False)]:
            db.session.add(QuestionOption(question_id=q.id, text=text, is_correct=correct, order=0))

        # Q6
        q = Question(assessment_id=a2.id, question_type='fill_blank',
                     text='Look! The children _____ (play) in the park.',
                     instruction='Type the correct form of the verb.', points=2, order=6)
        db.session.add(q)
        db.session.commit()
        db.session.add(QuestionOption(question_id=q.id, text='are playing', is_correct=True, order=0))

        # Q7
        q = Question(assessment_id=a2.id, question_type='matching',
                     text='Match the subject with the correct verb form.',
                     instruction='Match each subject with the correct present simple form.',
                     points=3, order=7)
        db.session.add(q)
        db.session.commit()
        matches = [('He', 'works'), ('They', 'work'), ('She', 'studies'), ('We', 'study')]
        for i, (item, match) in enumerate(matches):
            db.session.add(QuestionOption(question_id=q.id, text=item, match_text=match,
                                          is_correct=True, order=i))

        # Q8
        q = Question(assessment_id=a2.id, question_type='true_false',
                     text='The sentence "She play tennis every Sunday" is grammatically correct.',
                     instruction='Select True or False.', points=1, order=8)
        db.session.add(q)
        db.session.commit()
        db.session.add(QuestionOption(question_id=q.id, text='true', is_correct=False, order=0))
        db.session.add(QuestionOption(question_id=q.id, text='false', is_correct=True, order=1))

        db.session.commit()
        print(f'Created: {a2.title} (Code: {a2.access_code})')

        # ================================================================
        # ASSESSMENT 3: 3rd Grade Secondary - Pre-Intermediate (B1) - Reading
        # ================================================================
        grade_3s = Grade.query.filter_by(order=9).first()
        level_b1 = EnglishLevel.query.filter_by(code='B1').first()

        a3 = Assessment(
            title='Reading Comprehension: Daily Routines',
            description='Read the passage about daily routines and answer the questions.',
            assessment_type='homework',
            grade_id=grade_3s.id,
            english_level_id=level_b1.id,
            teacher_id=teacher.id,
            time_limit_minutes=0,
            show_results=True,
            access_code='READ01BB'
        )
        db.session.add(a3)
        db.session.commit()

        passage = (
            'Read the following passage:\n\n'
            '"Sarah is a 15-year-old student who lives in Lima, Peru. Every morning, she wakes up at '
            '6:00 AM and takes a shower. After getting dressed, she has breakfast with her family. She '
            'usually eats bread with butter and drinks a glass of orange juice. She walks to school '
            'because it is only ten minutes from her house. After school, she does her homework and '
            'then practices playing the guitar for one hour. In the evening, she helps her mother '
            'prepare dinner. She goes to bed at 10:00 PM."'
        )

        # Q1
        q = Question(assessment_id=a3.id, question_type='multiple_choice',
                     text=passage + '\n\nWhat time does Sarah wake up?',
                     instruction='Choose the correct answer based on the passage.',
                     points=2, order=1)
        db.session.add(q)
        db.session.commit()
        for text, correct in [('5:00 AM', False), ('6:00 AM', True), ('7:00 AM', False), ('8:00 AM', False)]:
            db.session.add(QuestionOption(question_id=q.id, text=text, is_correct=correct, order=0))

        # Q2
        q = Question(assessment_id=a3.id, question_type='multiple_choice',
                     text='How does Sarah go to school?',
                     instruction='Choose the correct answer.', points=2, order=2)
        db.session.add(q)
        db.session.commit()
        for text, correct in [('By bus', False), ('By car', False), ('She walks', True), ('By bicycle', False)]:
            db.session.add(QuestionOption(question_id=q.id, text=text, is_correct=correct, order=0))

        # Q3
        q = Question(assessment_id=a3.id, question_type='true_false',
                     text='Sarah drinks coffee for breakfast.',
                     instruction='Select True or False based on the passage.',
                     points=1, order=3)
        db.session.add(q)
        db.session.commit()
        db.session.add(QuestionOption(question_id=q.id, text='true', is_correct=False, order=0))
        db.session.add(QuestionOption(question_id=q.id, text='false', is_correct=True, order=1))

        # Q4
        q = Question(assessment_id=a3.id, question_type='fill_blank',
                     text='Sarah practices playing the _______ for one hour.',
                     instruction='Type the missing word from the passage.', points=2, order=4)
        db.session.add(q)
        db.session.commit()
        db.session.add(QuestionOption(question_id=q.id, text='guitar', is_correct=True, order=0))

        # Q5
        q = Question(assessment_id=a3.id, question_type='multiple_choice',
                     text='What does Sarah do in the evening?',
                     instruction='Choose the correct answer.', points=2, order=5)
        db.session.add(q)
        db.session.commit()
        for text, correct in [
            ('She watches TV', False),
            ('She helps her mother prepare dinner', True),
            ('She goes to the gym', False),
            ('She reads books', False)
        ]:
            db.session.add(QuestionOption(question_id=q.id, text=text, is_correct=correct, order=0))

        # Q6 - Short answer
        q = Question(assessment_id=a3.id, question_type='short_answer',
                     text='Write 3-4 sentences describing YOUR daily routine using present simple tense.',
                     instruction='Write in complete sentences. Use verbs like: wake up, eat, go, study, play.',
                     points=4, order=6)
        db.session.add(q)

        db.session.commit()
        print(f'Created: {a3.title} (Code: {a3.access_code})')

        # ================================================================
        # ASSESSMENT 4: 5th Grade Secondary - Intermediate (B2) - Advanced Grammar
        # ================================================================
        grade_5s = Grade.query.filter_by(order=11).first()
        level_b2 = EnglishLevel.query.filter_by(code='B2').first()

        a4 = Assessment(
            title='Conditionals and Reported Speech Exam',
            description='Comprehensive exam covering first, second, and third conditionals, as well as reported speech.',
            assessment_type='exam',
            grade_id=grade_5s.id,
            english_level_id=level_b2.id,
            teacher_id=teacher.id,
            time_limit_minutes=45,
            shuffle_questions=True,
            show_results=True,
            access_code='ADVGRAM01'
        )
        db.session.add(a4)
        db.session.commit()

        # Q1
        q = Question(assessment_id=a4.id, question_type='multiple_choice',
                     text='If I _____ more money, I would travel around the world.',
                     instruction='Choose the correct verb form (Second Conditional).', points=2, order=1)
        db.session.add(q)
        db.session.commit()
        for text, correct in [('have', False), ('had', True), ('will have', False), ('would have', False)]:
            db.session.add(QuestionOption(question_id=q.id, text=text, is_correct=correct, order=0))

        # Q2
        q = Question(assessment_id=a4.id, question_type='multiple_choice',
                     text='If it rains tomorrow, we _____ the picnic.',
                     instruction='Choose the correct verb form (First Conditional).', points=2, order=2)
        db.session.add(q)
        db.session.commit()
        for text, correct in [('cancel', False), ('will cancel', True), ('would cancel', False), ('cancelled', False)]:
            db.session.add(QuestionOption(question_id=q.id, text=text, is_correct=correct, order=0))

        # Q3
        q = Question(assessment_id=a4.id, question_type='multiple_choice',
                     text='If she had studied harder, she _____ the exam.',
                     instruction='Choose the correct verb form (Third Conditional).', points=2, order=3)
        db.session.add(q)
        db.session.commit()
        for text, correct in [
            ('would pass', False), ('would have passed', True),
            ('will pass', False), ('passed', False)
        ]:
            db.session.add(QuestionOption(question_id=q.id, text=text, is_correct=correct, order=0))

        # Q4
        q = Question(assessment_id=a4.id, question_type='fill_blank',
                     text='She said: "I am tired." -> She said that she _____ tired. (Reported Speech)',
                     instruction='Type the correct reported form.', points=2, order=4)
        db.session.add(q)
        db.session.commit()
        db.session.add(QuestionOption(question_id=q.id, text='was', is_correct=True, order=0))

        # Q5
        q = Question(assessment_id=a4.id, question_type='fill_blank',
                     text='He said: "I will call you tomorrow." -> He said that he _____ call me the next day.',
                     instruction='Type the correct word.', points=2, order=5)
        db.session.add(q)
        db.session.commit()
        db.session.add(QuestionOption(question_id=q.id, text='would', is_correct=True, order=0))

        # Q6
        q = Question(assessment_id=a4.id, question_type='true_false',
                     text='In the third conditional, we use "would have + past participle" in the result clause.',
                     instruction='Select True or False.', points=1, order=6)
        db.session.add(q)
        db.session.commit()
        db.session.add(QuestionOption(question_id=q.id, text='true', is_correct=True, order=0))
        db.session.add(QuestionOption(question_id=q.id, text='false', is_correct=False, order=1))

        # Q7
        q = Question(assessment_id=a4.id, question_type='matching',
                     text='Match each conditional type with its structure.',
                     instruction='Match the left column with the right column.',
                     points=3, order=7)
        db.session.add(q)
        db.session.commit()
        matches = [
            ('First Conditional', 'If + present simple, will + base verb'),
            ('Second Conditional', 'If + past simple, would + base verb'),
            ('Third Conditional', 'If + past perfect, would have + past participle'),
        ]
        for i, (item, match) in enumerate(matches):
            db.session.add(QuestionOption(question_id=q.id, text=item, match_text=match,
                                          is_correct=True, order=i))

        # Q8 - Short answer
        q = Question(assessment_id=a4.id, question_type='short_answer',
                     text='Write one sentence using each type of conditional (First, Second, and Third). '
                          'Label each sentence.',
                     instruction='Write 3 complete sentences, each clearly labeled.',
                     points=6, order=8)
        db.session.add(q)

        db.session.commit()
        print(f'Created: {a4.title} (Code: {a4.access_code})')

        # ================================================================
        # ASSESSMENT 5: 5th Grade Primary - Beginner (A1) - Practice
        # ================================================================
        grade_5p = Grade.query.filter_by(order=5).first()

        a5 = Assessment(
            title='Family Members and Body Parts Practice',
            description='Practice your vocabulary about family members and parts of the body.',
            assessment_type='practice',
            grade_id=grade_5p.id,
            english_level_id=level_a1.id,
            teacher_id=teacher.id,
            time_limit_minutes=0,
            show_results=True,
            access_code='FAMILY01'
        )
        db.session.add(a5)
        db.session.commit()

        # Q1
        q = Question(assessment_id=a5.id, question_type='multiple_choice',
                     text='Your mother\'s mother is your _____.',
                     instruction='Choose the correct answer.', points=2, order=1)
        db.session.add(q)
        db.session.commit()
        for text, correct in [('Aunt', False), ('Grandmother', True), ('Sister', False), ('Cousin', False)]:
            db.session.add(QuestionOption(question_id=q.id, text=text, is_correct=correct, order=0))

        # Q2
        q = Question(assessment_id=a5.id, question_type='matching',
                     text='Match each family member with the correct description.',
                     instruction='Match the items.', points=4, order=2)
        db.session.add(q)
        db.session.commit()
        matches = [
            ('Brother', 'Male sibling'),
            ('Uncle', "Parent's brother"),
            ('Cousin', "Aunt or uncle's child"),
            ('Nephew', "Sister or brother's son"),
        ]
        for i, (item, match) in enumerate(matches):
            db.session.add(QuestionOption(question_id=q.id, text=item, match_text=match,
                                          is_correct=True, order=i))

        # Q3
        q = Question(assessment_id=a5.id, question_type='fill_blank',
                     text='We use our _____ to see.',
                     instruction='Type the body part.', points=1, order=3)
        db.session.add(q)
        db.session.commit()
        db.session.add(QuestionOption(question_id=q.id, text='eyes', is_correct=True, order=0))

        # Q4
        q = Question(assessment_id=a5.id, question_type='fill_blank',
                     text='We use our _____ to hear sounds.',
                     instruction='Type the body part.', points=1, order=4)
        db.session.add(q)
        db.session.commit()
        db.session.add(QuestionOption(question_id=q.id, text='ears', is_correct=True, order=0))

        # Q5
        q = Question(assessment_id=a5.id, question_type='true_false',
                     text='Your father\'s sister is your aunt.',
                     instruction='Select True or False.', points=1, order=5)
        db.session.add(q)
        db.session.commit()
        db.session.add(QuestionOption(question_id=q.id, text='true', is_correct=True, order=0))
        db.session.add(QuestionOption(question_id=q.id, text='false', is_correct=False, order=1))

        db.session.commit()
        print(f'Created: {a5.title} (Code: {a5.access_code})')

        print('\n--- Seed complete! ---')
        print('Sample access codes:')
        print('  COLORS01  - 3rd Primary, Beginner (A1)')
        print('  GRAMMAR01 - 1st Secondary, Elementary (A2)')
        print('  READ01BB  - 3rd Secondary, Pre-Intermediate (B1)')
        print('  ADVGRAM01 - 5th Secondary, Intermediate (B2)')
        print('  FAMILY01  - 5th Primary, Beginner (A1)')


if __name__ == '__main__':
    seed_assessments()
