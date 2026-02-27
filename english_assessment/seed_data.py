"""
Seed data script for English Assessment Platform.
Populates the database with comprehensive exam models for all grades and levels.
Covers: 1ro-6to Primaria and 1ro-5to Secundaria.
Run: python seed_data.py
"""

import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

from app import app, db
from models import (Teacher, Grade, EnglishLevel, Assessment, Question,
                    QuestionOption)


def add_mc(assessment_id, text, instruction, points, order, options):
    """Helper: add a multiple choice question."""
    q = Question(assessment_id=assessment_id, question_type='multiple_choice',
                 text=text, instruction=instruction, points=points, order=order)
    db.session.add(q)
    db.session.flush()
    for i, (opt_text, correct) in enumerate(options):
        db.session.add(QuestionOption(question_id=q.id, text=opt_text,
                                      is_correct=correct, order=i))
    return q


def add_tf(assessment_id, text, instruction, points, order, correct_is_true):
    """Helper: add a true/false question."""
    q = Question(assessment_id=assessment_id, question_type='true_false',
                 text=text, instruction=instruction, points=points, order=order)
    db.session.add(q)
    db.session.flush()
    db.session.add(QuestionOption(question_id=q.id, text='true',
                                  is_correct=correct_is_true, order=0))
    db.session.add(QuestionOption(question_id=q.id, text='false',
                                  is_correct=not correct_is_true, order=1))
    return q


def add_fill(assessment_id, text, instruction, points, order, answer):
    """Helper: add a fill-in-the-blank question."""
    q = Question(assessment_id=assessment_id, question_type='fill_blank',
                 text=text, instruction=instruction, points=points, order=order)
    db.session.add(q)
    db.session.flush()
    db.session.add(QuestionOption(question_id=q.id, text=answer,
                                  is_correct=True, order=0))
    return q


def add_matching(assessment_id, text, instruction, points, order, pairs):
    """Helper: add a matching question."""
    q = Question(assessment_id=assessment_id, question_type='matching',
                 text=text, instruction=instruction, points=points, order=order)
    db.session.add(q)
    db.session.flush()
    for i, (left, right) in enumerate(pairs):
        db.session.add(QuestionOption(question_id=q.id, text=left,
                                      match_text=right, is_correct=True, order=i))
    return q


def add_ordering(assessment_id, text, instruction, points, order, items):
    """Helper: add an ordering question."""
    q = Question(assessment_id=assessment_id, question_type='ordering',
                 text=text, instruction=instruction, points=points, order=order)
    db.session.add(q)
    db.session.flush()
    for i, item in enumerate(items):
        db.session.add(QuestionOption(question_id=q.id, text=item,
                                      is_correct=True, order=i))
    return q


def add_short(assessment_id, text, instruction, points, order):
    """Helper: add a short answer question."""
    q = Question(assessment_id=assessment_id, question_type='short_answer',
                 text=text, instruction=instruction, points=points, order=order)
    db.session.add(q)
    db.session.flush()
    return q


def get_grade(order):
    return Grade.query.filter_by(order=order).first()


def get_level(code):
    return EnglishLevel.query.filter_by(code=code).first()


def seed_assessments():
    """Create comprehensive exam models for all grades and levels."""
    with app.app_context():
        if Assessment.query.first():
            print('Assessments already exist. Skipping seed.')
            print('To re-seed, delete the database file and run again.')
            return

        teacher = Teacher.query.filter_by(username='admin').first()
        if not teacher:
            print('No admin teacher found. Run the app first to initialize.')
            return

        tid = teacher.id
        all_codes = []

        # ==============================================================
        # PRIMARIA - 1ro a 6to Grado
        # ==============================================================

        # ---------------------------------------------------------------
        # 1. 1ro PRIMARIA - A1: Greetings, Numbers and Colors
        # ---------------------------------------------------------------
        a = Assessment(
            title='Greetings, Numbers and Colors',
            description='Learn and practice basic greetings, numbers 1-10, and primary colors in English.',
            assessment_type='exam',
            grade_id=get_grade(1).id,
            english_level_id=get_level('A1').id,
            teacher_id=tid,
            time_limit_minutes=15,
            show_results=True,
            access_code='GREET1PR'
        )
        db.session.add(a)
        db.session.flush()

        add_mc(a.id, 'What do you say when you meet someone in the morning?',
               'Choose the correct greeting.', 2, 1,
               [('Good morning!', True), ('Good night!', False),
                ('Goodbye!', False), ('Good evening!', False)])

        add_mc(a.id, 'How do you say "adiós" in English?',
               'Choose the correct answer.', 2, 2,
               [('Hello', False), ('Goodbye', True),
                ('Please', False), ('Thank you', False)])

        add_fill(a.id, 'Count: one, two, three, four, ______.',
                 'Type the next number in English.', 2, 3, 'five')

        add_fill(a.id, 'Count: six, seven, eight, nine, ______.',
                 'Type the next number in English.', 2, 4, 'ten')

        add_mc(a.id, 'What color is a banana?',
               'Choose the correct color.', 2, 5,
               [('Red', False), ('Blue', False),
                ('Yellow', True), ('Green', False)])

        add_tf(a.id, '"Good afternoon" means "Buenos días".',
               'Select True or False.', 1, 6, False)

        add_matching(a.id, 'Match the number with the word.',
                     'Match each number with its English name.', 3, 7,
                     [('3', 'three'), ('7', 'seven'), ('1', 'one')])

        add_mc(a.id, 'What color do you get when you mix red and blue?',
               'Choose the correct answer.', 2, 8,
               [('Green', False), ('Orange', False),
                ('Purple', True), ('Pink', False)])

        add_fill(a.id, 'When someone says "How are you?", you can answer: "I am ______, thank you."',
                 'Type one word (e.g., fine, good, happy).', 2, 9, 'fine')

        add_ordering(a.id, 'Put the numbers in order from smallest to largest.',
                     'Drag and arrange in the correct order.', 3, 10,
                     ['two', 'five', 'eight', 'ten'])

        db.session.commit()
        all_codes.append(('GREET1PR', '1st Primary', 'Beginner (A1)', a.title))
        print(f'  Created: {a.title}')

        # ---------------------------------------------------------------
        # 2. 1ro PRIMARIA - A1: Classroom Objects
        # ---------------------------------------------------------------
        a = Assessment(
            title='My Classroom Objects',
            description='Identify and name common objects found in the classroom.',
            assessment_type='quiz',
            grade_id=get_grade(1).id,
            english_level_id=get_level('A1').id,
            teacher_id=tid,
            time_limit_minutes=10,
            show_results=True,
            access_code='CLASS1PR'
        )
        db.session.add(a)
        db.session.flush()

        add_mc(a.id, 'You write with a _____.',
               'Choose the correct classroom object.', 2, 1,
               [('Book', False), ('Pencil', True), ('Eraser', False), ('Ruler', False)])

        add_mc(a.id, 'You cut paper with _____.',
               'Choose the correct answer.', 2, 2,
               [('Scissors', True), ('Glue', False), ('Crayon', False), ('Notebook', False)])

        add_fill(a.id, 'I read stories from a ______.',
                 'Type the classroom object.', 2, 3, 'book')

        add_tf(a.id, 'A ruler is used to draw straight lines.',
               'Select True or False.', 1, 4, True)

        add_matching(a.id, 'Match the object with what you do with it.',
                     'Match correctly.', 4, 5,
                     [('Pencil', 'Write'), ('Eraser', 'Erase'),
                      ('Scissors', 'Cut'), ('Glue', 'Stick')])

        add_mc(a.id, 'Where do you keep your pencils?',
               'Choose the correct answer.', 2, 6,
               [('In a pencil case', True), ('In a plate', False),
                ('In a cup', False), ('In a shoe', False)])

        db.session.commit()
        all_codes.append(('CLASS1PR', '1st Primary', 'Beginner (A1)', a.title))
        print(f'  Created: {a.title}')

        # ---------------------------------------------------------------
        # 3. 2do PRIMARIA - A1: Fruits, Food and Drinks
        # ---------------------------------------------------------------
        a = Assessment(
            title='Fruits, Food and Drinks Vocabulary',
            description='Learn the names of common fruits, foods, and drinks in English.',
            assessment_type='exam',
            grade_id=get_grade(2).id,
            english_level_id=get_level('A1').id,
            teacher_id=tid,
            time_limit_minutes=15,
            show_results=True,
            access_code='FOOD2PRI'
        )
        db.session.add(a)
        db.session.flush()

        add_mc(a.id, 'Which one is a fruit?',
               'Choose the correct answer.', 2, 1,
               [('Bread', False), ('Apple', True), ('Rice', False), ('Milk', False)])

        add_mc(a.id, 'What do you drink for breakfast?',
               'Choose the most common answer.', 2, 2,
               [('Soup', False), ('Juice', True), ('Cake', False), ('Salad', False)])

        add_fill(a.id, 'A ______ is a yellow fruit that monkeys like.',
                 'Type the name of the fruit.', 2, 3, 'banana')

        add_tf(a.id, 'Water is a type of food.',
               'Select True or False.', 1, 4, False)

        add_matching(a.id, 'Match each food item with its category.',
                     'Match the item on the left with its type on the right.', 4, 5,
                     [('Apple', 'Fruit'), ('Milk', 'Drink'),
                      ('Bread', 'Food'), ('Orange juice', 'Drink')])

        add_mc(a.id, 'Which of these is NOT a vegetable?',
               'Choose the correct answer.', 2, 6,
               [('Carrot', False), ('Tomato', False),
                ('Strawberry', True), ('Potato', False)])

        add_fill(a.id, 'Ice ______ is a cold and sweet dessert.',
                 'Type the missing word.', 2, 7, 'cream')

        add_mc(a.id, 'What color is a watermelon inside?',
               'Choose the correct answer.', 2, 8,
               [('Yellow', False), ('Green', False),
                ('Red', True), ('Blue', False)])

        add_ordering(a.id, 'Put these meals in order during the day.',
                     'Arrange from first to last meal.', 3, 9,
                     ['Breakfast', 'Lunch', 'Snack', 'Dinner'])

        add_tf(a.id, 'Cheese is made from milk.',
               'Select True or False.', 1, 10, True)

        db.session.commit()
        all_codes.append(('FOOD2PRI', '2nd Primary', 'Beginner (A1)', a.title))
        print(f'  Created: {a.title}')

        # ---------------------------------------------------------------
        # 4. 2do PRIMARIA - A1: My Body Parts
        # ---------------------------------------------------------------
        a = Assessment(
            title='My Body Parts',
            description='Identify and name the main parts of the human body.',
            assessment_type='quiz',
            grade_id=get_grade(2).id,
            english_level_id=get_level('A1').id,
            teacher_id=tid,
            time_limit_minutes=10,
            show_results=True,
            access_code='BODY2PRI'
        )
        db.session.add(a)
        db.session.flush()

        add_mc(a.id, 'You use your _____ to walk.',
               'Choose the correct body part.', 2, 1,
               [('Hands', False), ('Legs', True), ('Arms', False), ('Head', False)])

        add_fill(a.id, 'You use your ______ to think.',
                 'Type the body part.', 2, 2, 'brain')

        add_mc(a.id, 'How many fingers do you have on one hand?',
               'Choose the correct number.', 2, 3,
               [('Three', False), ('Four', False), ('Five', True), ('Six', False)])

        add_matching(a.id, 'Match the body part with its function.',
                     'Match each body part with what it does.', 4, 4,
                     [('Eyes', 'See'), ('Ears', 'Hear'),
                      ('Nose', 'Smell'), ('Mouth', 'Eat and talk')])

        add_tf(a.id, 'Your elbow is part of your leg.',
               'Select True or False.', 1, 5, False)

        add_fill(a.id, 'You have ten ______ on your feet.',
                 'Type the body part.', 2, 6, 'toes')

        db.session.commit()
        all_codes.append(('BODY2PRI', '2nd Primary', 'Beginner (A1)', a.title))
        print(f'  Created: {a.title}')

        # ---------------------------------------------------------------
        # 5. 3ro PRIMARIA - A1: Colors and Animals Vocabulary Quiz
        # ---------------------------------------------------------------
        a = Assessment(
            title='Colors and Animals Vocabulary Quiz',
            description='Test your knowledge of basic colors and animal names in English.',
            assessment_type='quiz',
            grade_id=get_grade(3).id,
            english_level_id=get_level('A1').id,
            teacher_id=tid,
            time_limit_minutes=15,
            show_results=True,
            access_code='COLORS01'
        )
        db.session.add(a)
        db.session.flush()

        add_mc(a.id, 'What color is the sky on a clear day?',
               'Choose the correct answer.', 2, 1,
               [('Red', False), ('Blue', True), ('Green', False), ('Yellow', False)])

        add_mc(a.id, 'Which animal says "meow"?',
               'Choose the correct answer.', 2, 2,
               [('Dog', False), ('Cat', True), ('Bird', False), ('Fish', False)])

        add_tf(a.id, 'A dog is a pet animal.',
               'Select True or False.', 1, 3, True)

        add_fill(a.id, 'The color of grass is ______.',
                 'Type the missing word.', 2, 4, 'green')

        add_matching(a.id, 'Match each animal with its sound.',
                     'Match the items on the left with the correct option on the right.',
                     3, 5,
                     [('Dog', 'Woof'), ('Cat', 'Meow'), ('Cow', 'Moo')])

        add_mc(a.id, 'How many legs does a spider have?',
               'Choose the correct answer.', 2, 6,
               [('4', False), ('6', False), ('8', True), ('10', False)])

        db.session.commit()
        all_codes.append(('COLORS01', '3rd Primary', 'Beginner (A1)', a.title))
        print(f'  Created: {a.title}')

        # ---------------------------------------------------------------
        # 6. 3ro PRIMARIA - A1: Days, Months and Seasons
        # ---------------------------------------------------------------
        a = Assessment(
            title='Days of the Week, Months and Seasons',
            description='Practice the days of the week, months of the year, and the four seasons.',
            assessment_type='exam',
            grade_id=get_grade(3).id,
            english_level_id=get_level('A1').id,
            teacher_id=tid,
            time_limit_minutes=20,
            show_results=True,
            access_code='DAYS3PRI'
        )
        db.session.add(a)
        db.session.flush()

        add_mc(a.id, 'How many days are there in a week?',
               'Choose the correct answer.', 2, 1,
               [('5', False), ('6', False), ('7', True), ('8', False)])

        add_fill(a.id, 'The day after Monday is ______.',
                 'Type the correct day.', 2, 2, 'Tuesday')

        add_mc(a.id, 'Which month comes after March?',
               'Choose the correct answer.', 2, 3,
               [('February', False), ('April', True), ('May', False), ('January', False)])

        add_ordering(a.id, 'Put the days of the week in order.',
                     'Arrange from the first day to the last.', 4, 4,
                     ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'])

        add_matching(a.id, 'Match the season with the weather.',
                     'Match each season with the correct description.', 4, 5,
                     [('Summer', 'Hot and sunny'), ('Winter', 'Cold and snowy'),
                      ('Spring', 'Warm and rainy'), ('Autumn', 'Cool and windy')])

        add_tf(a.id, 'December is a summer month.',
               'Select True or False. (Think about Peru!)', 1, 6, True)

        add_mc(a.id, 'Which month has 28 (or 29) days?',
               'Choose the correct answer.', 2, 7,
               [('January', False), ('February', True), ('March', False), ('April', False)])

        add_fill(a.id, 'The weekend days are Saturday and ______.',
                 'Type the missing day.', 2, 8, 'Sunday')

        add_mc(a.id, 'In which season do leaves fall from the trees?',
               'Choose the correct season.', 2, 9,
               [('Spring', False), ('Summer', False), ('Autumn', True), ('Winter', False)])

        add_ordering(a.id, 'Put these months in order.',
                     'Arrange from first to last.', 3, 10,
                     ['January', 'April', 'July', 'October'])

        db.session.commit()
        all_codes.append(('DAYS3PRI', '3rd Primary', 'Beginner (A1)', a.title))
        print(f'  Created: {a.title}')

        # ---------------------------------------------------------------
        # 7. 4to PRIMARIA - A1: Clothes and Weather
        # ---------------------------------------------------------------
        a = Assessment(
            title='Clothes and Weather Vocabulary',
            description='Learn vocabulary about clothing and weather conditions. Describe what to wear in different weather.',
            assessment_type='exam',
            grade_id=get_grade(4).id,
            english_level_id=get_level('A1').id,
            teacher_id=tid,
            time_limit_minutes=20,
            show_results=True,
            access_code='CLTH4PRI'
        )
        db.session.add(a)
        db.session.flush()

        add_mc(a.id, 'What do you wear on your feet?',
               'Choose the correct answer.', 2, 1,
               [('Hat', False), ('Shoes', True), ('Gloves', False), ('Scarf', False)])

        add_mc(a.id, 'When it is raining, you should carry an _____.',
               'Choose the correct answer.', 2, 2,
               [('Umbrella', True), ('Sunglasses', False),
                ('Ice cream', False), ('Fan', False)])

        add_fill(a.id, 'When it is cold, you wear a ______ to stay warm.',
                 'Type a piece of clothing.', 2, 3, 'jacket')

        add_matching(a.id, 'Match the weather with the correct clothing.',
                     'Match each weather condition with what you should wear.', 4, 4,
                     [('Sunny', 'Sunglasses'), ('Rainy', 'Raincoat'),
                      ('Cold', 'Scarf'), ('Hot', 'T-shirt')])

        add_tf(a.id, 'You wear a swimsuit when it is snowing.',
               'Select True or False.', 1, 5, False)

        add_mc(a.id, 'What is the weather like when there are dark clouds?',
               'Choose the correct answer.', 2, 6,
               [('Sunny', False), ('Cloudy and rainy', True),
                ('Windy', False), ('Snowy', False)])

        add_fill(a.id, 'In summer, the weather is usually ______ and sunny.',
                 'Type the missing word.', 2, 7, 'hot')

        add_mc(a.id, 'Which of these is NOT a piece of clothing?',
               'Choose the correct answer.', 2, 8,
               [('Dress', False), ('Pants', False),
                ('Cloud', True), ('Socks', False)])

        add_ordering(a.id, 'Put the clothes in order from bottom to top (as you wear them).',
                     'Arrange from what goes on first.', 3, 9,
                     ['Socks', 'Pants', 'Shirt', 'Hat'])

        add_short(a.id, 'Look outside. What is the weather like today? What are you wearing? Write 2-3 sentences.',
                  'Use words like: sunny, cloudy, hot, cold, wearing, shirt, pants, shoes.', 4, 10)

        db.session.commit()
        all_codes.append(('CLTH4PRI', '4th Primary', 'Beginner (A1)', a.title))
        print(f'  Created: {a.title}')

        # ---------------------------------------------------------------
        # 8. 4to PRIMARIA - A1: My House and Furniture
        # ---------------------------------------------------------------
        a = Assessment(
            title='My House: Rooms and Furniture',
            description='Learn the names of rooms in a house and common furniture.',
            assessment_type='quiz',
            grade_id=get_grade(4).id,
            english_level_id=get_level('A1').id,
            teacher_id=tid,
            time_limit_minutes=15,
            show_results=True,
            access_code='HOUS4PRI'
        )
        db.session.add(a)
        db.session.flush()

        add_mc(a.id, 'Where do you cook food?',
               'Choose the correct room.', 2, 1,
               [('Bedroom', False), ('Kitchen', True),
                ('Bathroom', False), ('Living room', False)])

        add_mc(a.id, 'Where do you sleep?',
               'Choose the correct room.', 2, 2,
               [('Kitchen', False), ('Garden', False),
                ('Bedroom', True), ('Dining room', False)])

        add_fill(a.id, 'You watch TV in the ______ room.',
                 'Type the missing word.', 2, 3, 'living')

        add_matching(a.id, 'Match the furniture with the correct room.',
                     'Match each piece of furniture with the room where it usually is.', 4, 4,
                     [('Bed', 'Bedroom'), ('Stove', 'Kitchen'),
                      ('Sofa', 'Living room'), ('Toilet', 'Bathroom')])

        add_tf(a.id, 'The refrigerator is usually in the bedroom.',
               'Select True or False.', 1, 5, False)

        add_mc(a.id, 'What furniture do you sit on to eat dinner?',
               'Choose the correct answer.', 2, 6,
               [('Bed', False), ('Chair', True), ('Bathtub', False), ('Wardrobe', False)])

        add_fill(a.id, 'You wash your hands in the ______.',
                 'Type the room name.', 2, 7, 'bathroom')

        add_tf(a.id, 'A garden is a part of many houses where plants grow.',
               'Select True or False.', 1, 8, True)

        db.session.commit()
        all_codes.append(('HOUS4PRI', '4th Primary', 'Beginner (A1)', a.title))
        print(f'  Created: {a.title}')

        # ---------------------------------------------------------------
        # 9. 5to PRIMARIA - A1: Family Members and Body Parts Practice
        # ---------------------------------------------------------------
        a = Assessment(
            title='Family Members and Body Parts Practice',
            description='Practice your vocabulary about family members and parts of the body.',
            assessment_type='practice',
            grade_id=get_grade(5).id,
            english_level_id=get_level('A1').id,
            teacher_id=tid,
            time_limit_minutes=0,
            show_results=True,
            access_code='FAMILY01'
        )
        db.session.add(a)
        db.session.flush()

        add_mc(a.id, 'Your mother\'s mother is your _____.',
               'Choose the correct answer.', 2, 1,
               [('Aunt', False), ('Grandmother', True), ('Sister', False), ('Cousin', False)])

        add_matching(a.id, 'Match each family member with the correct description.',
                     'Match the items.', 4, 2,
                     [('Brother', 'Male sibling'), ('Uncle', "Parent's brother"),
                      ('Cousin', "Aunt or uncle's child"), ('Nephew', "Sister or brother's son")])

        add_fill(a.id, 'We use our _____ to see.',
                 'Type the body part.', 1, 3, 'eyes')

        add_fill(a.id, 'We use our _____ to hear sounds.',
                 'Type the body part.', 1, 4, 'ears')

        add_tf(a.id, 'Your father\'s sister is your aunt.',
               'Select True or False.', 1, 5, True)

        db.session.commit()
        all_codes.append(('FAMILY01', '5th Primary', 'Beginner (A1)', a.title))
        print(f'  Created: {a.title}')

        # ---------------------------------------------------------------
        # 10. 5to PRIMARIA - A2: Prepositions of Place and There is/are
        # ---------------------------------------------------------------
        a = Assessment(
            title='Prepositions of Place and There is / There are',
            description='Practice using prepositions of place (in, on, under, next to, between) and "there is / there are" structures.',
            assessment_type='exam',
            grade_id=get_grade(5).id,
            english_level_id=get_level('A2').id,
            teacher_id=tid,
            time_limit_minutes=25,
            show_results=True,
            access_code='PREP5PRI'
        )
        db.session.add(a)
        db.session.flush()

        add_mc(a.id, 'The cat is _____ the table.',
               'Choose the correct preposition (the cat is below the table).', 2, 1,
               [('on', False), ('in', False), ('under', True), ('next to', False)])

        add_mc(a.id, '_____ a book on the desk.',
               'Choose the correct form.', 2, 2,
               [('There is', True), ('There are', False),
                ('There was', False), ('There were', False)])

        add_fill(a.id, '_____ three apples in the basket.',
                 'Type "There is" or "There are".', 2, 3, 'There are')

        add_mc(a.id, 'The lamp is _____ the table.',
               'Choose the correct preposition (the lamp is above the table surface).', 2, 4,
               [('under', False), ('on', True), ('in', False), ('behind', False)])

        add_matching(a.id, 'Match the preposition with its meaning.',
                     'Match each English preposition with the Spanish translation.', 4, 5,
                     [('On', 'Sobre/encima de'), ('Under', 'Debajo de'),
                      ('Next to', 'Al lado de'), ('Between', 'Entre')])

        add_tf(a.id, '"There are a cat on the sofa" is grammatically correct.',
               'Select True or False.', 1, 6, False)

        add_fill(a.id, 'The school is _____ the park and the hospital.',
                 'Type the preposition that means "in the middle of two things".', 2, 7, 'between')

        add_mc(a.id, '_____ any milk in the fridge?',
               'Choose the correct form.', 2, 8,
               [('Is there', True), ('Are there', False),
                ('There is', False), ('There are', False)])

        add_short(a.id, 'Describe your bedroom. Where are 3 objects? Use prepositions of place.',
                  'Example: "There is a lamp on my desk. My shoes are under the bed." Write at least 3 sentences.', 4, 9)

        add_tf(a.id, '"There is" is used for singular nouns and "there are" for plural nouns.',
               'Select True or False.', 1, 10, True)

        db.session.commit()
        all_codes.append(('PREP5PRI', '5th Primary', 'Elementary (A2)', a.title))
        print(f'  Created: {a.title}')

        # ---------------------------------------------------------------
        # 11. 6to PRIMARIA - A2: Past Simple - Regular and Irregular Verbs
        # ---------------------------------------------------------------
        a = Assessment(
            title='Past Simple: Regular and Irregular Verbs',
            description='Learn and practice the past simple tense with regular (-ed) and common irregular verbs.',
            assessment_type='exam',
            grade_id=get_grade(6).id,
            english_level_id=get_level('A2').id,
            teacher_id=tid,
            time_limit_minutes=30,
            show_results=True,
            access_code='PAST6PRI'
        )
        db.session.add(a)
        db.session.flush()

        add_mc(a.id, 'Yesterday I _____ to the park. (go)',
               'Choose the correct past form.', 2, 1,
               [('go', False), ('goed', False), ('went', True), ('goes', False)])

        add_mc(a.id, 'She _____ a delicious cake last weekend. (make)',
               'Choose the correct past form.', 2, 2,
               [('maked', False), ('made', True), ('makes', False), ('making', False)])

        add_fill(a.id, 'We ______ (play) football yesterday afternoon.',
                 'Type the past simple form of the verb.', 2, 3, 'played')

        add_fill(a.id, 'He ______ (eat) pizza for dinner last night.',
                 'Type the past simple form of the verb.', 2, 4, 'ate')

        add_tf(a.id, 'The past simple of "study" is "studyed".',
               'Select True or False.', 1, 5, False)

        add_matching(a.id, 'Match each verb with its past simple form.',
                     'Match the present form with its correct past form.', 4, 6,
                     [('Go', 'Went'), ('Have', 'Had'),
                      ('See', 'Saw'), ('Come', 'Came')])

        add_mc(a.id, '_____ you watch TV last night?',
               'Choose the correct auxiliary for past simple questions.', 2, 7,
               [('Do', False), ('Did', True), ('Does', False), ('Was', False)])

        add_fill(a.id, 'They ______ (not/go) to school yesterday.',
                 'Type the correct negative past form.', 2, 8, "didn't go")

        add_mc(a.id, 'I _____ my homework before dinner. (finish)',
               'Choose the correct past form.', 2, 9,
               [('finish', False), ('finishes', False),
                ('finishing', False), ('finished', True)])

        add_ordering(a.id, 'Put the words in order to make a sentence.',
                     'Arrange to form: "She went to the cinema yesterday."', 3, 10,
                     ['She', 'went', 'to the cinema', 'yesterday'])

        add_matching(a.id, 'Match the regular verb with its past form.',
                     'Match each verb with its -ed form.', 3, 11,
                     [('Walk', 'Walked'), ('Dance', 'Danced'), ('Stop', 'Stopped')])

        add_short(a.id, 'Write about what you did last weekend. Use at least 4 past simple verbs.',
                  'Example: "Last Saturday I played football with my friends. We ate ice cream." Write 3-4 sentences.', 4, 12)

        db.session.commit()
        all_codes.append(('PAST6PRI', '6th Primary', 'Elementary (A2)', a.title))
        print(f'  Created: {a.title}')

        # ---------------------------------------------------------------
        # 12. 6to PRIMARIA - A2: Can / Can't - Abilities
        # ---------------------------------------------------------------
        a = Assessment(
            title='Can and Can\'t: Talking About Abilities',
            description='Practice using "can" and "can\'t" to talk about abilities and ask for permission.',
            assessment_type='quiz',
            grade_id=get_grade(6).id,
            english_level_id=get_level('A2').id,
            teacher_id=tid,
            time_limit_minutes=15,
            show_results=True,
            access_code='ABIL6PRI'
        )
        db.session.add(a)
        db.session.flush()

        add_mc(a.id, 'Fish _____ swim very well.',
               'Choose the correct answer.', 2, 1,
               [('can', True), ("can't", False), ('do', False), ('are', False)])

        add_mc(a.id, 'Birds _____ fly, but penguins _____.',
               'Choose the correct pair.', 2, 2,
               [("can / can't", True), ("can't / can", False),
                ("can / can", False), ("can't / can't", False)])

        add_fill(a.id, '______ you play the guitar? Yes, I can!',
                 'Type the missing word.', 2, 3, 'Can')

        add_tf(a.id, '"Can" is used to talk about ability in the present.',
               'Select True or False.', 1, 4, True)

        add_matching(a.id, 'Match the animal with its ability.',
                     'Match each animal with what it CAN do.', 3, 5,
                     [('Eagle', 'Can fly very high'), ('Cheetah', 'Can run very fast'),
                      ('Dolphin', 'Can swim and jump')])

        add_mc(a.id, 'My little brother is 2 years old. He _____ read books.',
               'Choose the correct answer.', 2, 6,
               [('can', False), ("can't", True), ('does', False), ("don't", False)])

        add_short(a.id, 'Write 4 sentences about what you can and can\'t do.',
                  'Example: "I can play football. I can\'t speak French." Write 4 sentences: 2 with CAN, 2 with CAN\'T.',
                  4, 7)

        db.session.commit()
        all_codes.append(('ABIL6PRI', '6th Primary', 'Elementary (A2)', a.title))
        print(f'  Created: {a.title}')

        # ==============================================================
        # SECUNDARIA - 1ro a 5to Grado
        # ==============================================================

        # ---------------------------------------------------------------
        # 13. 1ro SECUNDARIA - A2: Present Simple and Continuous Grammar
        # ---------------------------------------------------------------
        a = Assessment(
            title='Present Simple and Continuous Grammar Test',
            description='Evaluate your understanding of present simple and present continuous tenses.',
            assessment_type='exam',
            grade_id=get_grade(7).id,
            english_level_id=get_level('A2').id,
            teacher_id=tid,
            time_limit_minutes=30,
            show_results=True,
            access_code='GRAMMAR01'
        )
        db.session.add(a)
        db.session.flush()

        add_mc(a.id, 'She _____ to school every day.',
               'Choose the correct form of the verb.', 2, 1,
               [('go', False), ('goes', True), ('going', False), ('is go', False)])

        add_mc(a.id, 'They _____ football right now.',
               'Choose the correct form of the verb.', 2, 2,
               [('play', False), ('plays', False), ('are playing', True), ('is playing', False)])

        add_tf(a.id, 'We use present continuous for actions happening right now.',
               'Select True or False.', 1, 3, True)

        add_fill(a.id, 'He _____ (not/like) coffee. (Use present simple negative)',
                 'Type the correct form.', 2, 4, "doesn't like")

        add_mc(a.id, '_____ you usually wake up early?',
               'Choose the correct auxiliary verb.', 2, 5,
               [('Do', True), ('Does', False), ('Are', False), ('Is', False)])

        add_fill(a.id, 'Look! The children _____ (play) in the park.',
                 'Type the correct form of the verb.', 2, 6, 'are playing')

        add_matching(a.id, 'Match the subject with the correct verb form.',
                     'Match each subject with the correct present simple form.',
                     3, 7,
                     [('He', 'works'), ('They', 'work'), ('She', 'studies'), ('We', 'study')])

        add_tf(a.id, 'The sentence "She play tennis every Sunday" is grammatically correct.',
               'Select True or False.', 1, 8, False)

        db.session.commit()
        all_codes.append(('GRAMMAR01', '1st Secondary', 'Elementary (A2)', a.title))
        print(f'  Created: {a.title}')

        # ---------------------------------------------------------------
        # 14. 1ro SECUNDARIA - A2: Describing People and Adjectives
        # ---------------------------------------------------------------
        a = Assessment(
            title='Describing People: Adjectives and Physical Appearance',
            description='Practice describing people using adjectives for personality and physical appearance.',
            assessment_type='quiz',
            grade_id=get_grade(7).id,
            english_level_id=get_level('A2').id,
            teacher_id=tid,
            time_limit_minutes=20,
            show_results=True,
            access_code='DESC1SEC'
        )
        db.session.add(a)
        db.session.flush()

        add_mc(a.id, 'A person who helps others is _____.',
               'Choose the best adjective.', 2, 1,
               [('Lazy', False), ('Kind', True), ('Shy', False), ('Angry', False)])

        add_mc(a.id, 'She has long, _____ hair and blue eyes.',
               'Choose the best description.', 2, 2,
               [('short', False), ('blonde', True), ('bald', False), ('curly', False)])

        add_fill(a.id, 'The opposite of "tall" is ______.',
                 'Type the opposite adjective.', 2, 3, 'short')

        add_matching(a.id, 'Match each adjective with its opposite.',
                     'Match the adjectives.', 4, 4,
                     [('Happy', 'Sad'), ('Tall', 'Short'),
                      ('Fat', 'Thin'), ('Young', 'Old')])

        add_tf(a.id, '"Handsome" is usually used to describe a good-looking man.',
               'Select True or False.', 1, 5, True)

        add_mc(a.id, 'He is very _____. He always makes everyone laugh.',
               'Choose the best personality adjective.', 2, 6,
               [('Boring', False), ('Funny', True), ('Quiet', False), ('Serious', False)])

        add_fill(a.id, 'My grandmother is very ______. She always tells interesting stories. (wise/smart)',
                 'Type an adjective that describes intelligence.', 2, 7, 'wise')

        add_short(a.id, 'Describe your best friend. Write about their appearance and personality (4-5 sentences).',
                  'Use adjectives: tall, short, funny, kind, smart, etc. Use "He/She has..." and "He/She is..."',
                  4, 8)

        db.session.commit()
        all_codes.append(('DESC1SEC', '1st Secondary', 'Elementary (A2)', a.title))
        print(f'  Created: {a.title}')

        # ---------------------------------------------------------------
        # 15. 2do SECUNDARIA - A2: Comparative and Superlative Adjectives
        # ---------------------------------------------------------------
        a = Assessment(
            title='Comparative and Superlative Adjectives',
            description='Practice forming and using comparative and superlative forms of adjectives.',
            assessment_type='exam',
            grade_id=get_grade(8).id,
            english_level_id=get_level('A2').id,
            teacher_id=tid,
            time_limit_minutes=30,
            show_results=True,
            access_code='COMP2SEC'
        )
        db.session.add(a)
        db.session.flush()

        add_mc(a.id, 'My house is _____ than your house.',
               'Choose the correct comparative form of "big".', 2, 1,
               [('biger', False), ('bigger', True), ('more big', False), ('biggest', False)])

        add_mc(a.id, 'She is the _____ student in the class.',
               'Choose the correct superlative form of "intelligent".', 2, 2,
               [('intelligentest', False), ('more intelligent', False),
                ('most intelligent', True), ('intelligenter', False)])

        add_fill(a.id, 'A car is ______ than a bicycle. (fast)',
                 'Type the comparative form.', 2, 3, 'faster')

        add_fill(a.id, 'Mount Everest is the ______ mountain in the world. (high)',
                 'Type the superlative form.', 2, 4, 'highest')

        add_matching(a.id, 'Match each adjective with its comparative form.',
                     'Match correctly.', 4, 5,
                     [('Good', 'Better'), ('Bad', 'Worse'),
                      ('Far', 'Farther'), ('Little', 'Less')])

        add_tf(a.id, 'For short adjectives (one syllable), we add -er for comparative and -est for superlative.',
               'Select True or False.', 1, 6, True)

        add_mc(a.id, 'English is _____ interesting _____ math. (I prefer English)',
               'Choose the correct comparative structure.', 2, 7,
               [('more / than', True), ('most / than', False),
                ('more / that', False), ('most / of', False)])

        add_tf(a.id, '"More better" is a correct comparative form.',
               'Select True or False.', 1, 8, False)

        add_mc(a.id, 'This is the _____ movie I have ever seen!',
               'Choose the correct superlative of "bad".', 2, 9,
               [('baddest', False), ('worst', True), ('most bad', False), ('worse', False)])

        add_ordering(a.id, 'Order from smallest to biggest.',
                     'Arrange the animals by size.', 3, 10,
                     ['Mouse', 'Cat', 'Dog', 'Elephant'])

        add_short(a.id, 'Compare two cities you know. Write 4 sentences using comparatives and superlatives.',
                  'Example: "Lima is bigger than Piura. It is the largest city in Peru." Use: bigger, smaller, more beautiful, the best, etc.',
                  4, 11)

        db.session.commit()
        all_codes.append(('COMP2SEC', '2nd Secondary', 'Elementary (A2)', a.title))
        print(f'  Created: {a.title}')

        # ---------------------------------------------------------------
        # 16. 2do SECUNDARIA - B1: Present Perfect Tense
        # ---------------------------------------------------------------
        a = Assessment(
            title='Present Perfect Tense: Have/Has + Past Participle',
            description='Practice the present perfect tense for experiences, recent actions, and unfinished time periods.',
            assessment_type='exam',
            grade_id=get_grade(8).id,
            english_level_id=get_level('B1').id,
            teacher_id=tid,
            time_limit_minutes=30,
            show_results=True,
            access_code='PERF2SEC'
        )
        db.session.add(a)
        db.session.flush()

        add_mc(a.id, 'I _____ never been to Europe.',
               'Choose the correct auxiliary.', 2, 1,
               [('has', False), ('have', True), ('had', False), ('am', False)])

        add_mc(a.id, 'She _____ already finished her homework.',
               'Choose the correct auxiliary.', 2, 2,
               [('have', False), ('has', True), ('is', False), ('was', False)])

        add_fill(a.id, 'They ______ (visit) Paris three times.',
                 'Type the present perfect form.', 2, 3, 'have visited')

        add_fill(a.id, 'He ______ (lose) his keys. He can\'t find them.',
                 'Type the present perfect form.', 2, 4, 'has lost')

        add_mc(a.id, '_____ you ever eaten sushi?',
               'Choose the correct form.', 2, 5,
               [('Have', True), ('Has', False), ('Did', False), ('Do', False)])

        add_matching(a.id, 'Match the verb with its past participle.',
                     'Match each base verb with its past participle form.', 4, 6,
                     [('Write', 'Written'), ('Break', 'Broken'),
                      ('Speak', 'Spoken'), ('Drive', 'Driven')])

        add_tf(a.id, 'We use "already" in present perfect to say something happened sooner than expected.',
               'Select True or False.', 1, 7, True)

        add_mc(a.id, 'I haven\'t seen that movie _____.',
               'Choose the correct word.', 2, 8,
               [('already', False), ('yet', True), ('just', False), ('ever', False)])

        add_tf(a.id, '"I have went to the store" is grammatically correct.',
               'Select True or False.', 1, 9, False)

        add_short(a.id, 'Write 4 sentences about things you have done (or have never done) in your life using present perfect.',
                  'Use: have been, have eaten, have visited, have never tried, etc.',
                  4, 10)

        db.session.commit()
        all_codes.append(('PERF2SEC', '2nd Secondary', 'Pre-Intermediate (B1)', a.title))
        print(f'  Created: {a.title}')

        # ---------------------------------------------------------------
        # 17. 3ro SECUNDARIA - B1: Reading Comprehension: Daily Routines
        # ---------------------------------------------------------------
        a = Assessment(
            title='Reading Comprehension: Daily Routines',
            description='Read the passage about daily routines and answer the questions.',
            assessment_type='homework',
            grade_id=get_grade(9).id,
            english_level_id=get_level('B1').id,
            teacher_id=tid,
            time_limit_minutes=0,
            show_results=True,
            access_code='READ01BB'
        )
        db.session.add(a)
        db.session.flush()

        passage = (
            'Read the following passage:\n\n'
            '"Sarah is a 15-year-old student who lives in Lima, Peru. Every morning, she wakes up at '
            '6:00 AM and takes a shower. After getting dressed, she has breakfast with her family. She '
            'usually eats bread with butter and drinks a glass of orange juice. She walks to school '
            'because it is only ten minutes from her house. After school, she does her homework and '
            'then practices playing the guitar for one hour. In the evening, she helps her mother '
            'prepare dinner. She goes to bed at 10:00 PM."'
        )

        add_mc(a.id, passage + '\n\nWhat time does Sarah wake up?',
               'Choose the correct answer based on the passage.', 2, 1,
               [('5:00 AM', False), ('6:00 AM', True), ('7:00 AM', False), ('8:00 AM', False)])

        add_mc(a.id, 'How does Sarah go to school?',
               'Choose the correct answer.', 2, 2,
               [('By bus', False), ('By car', False), ('She walks', True), ('By bicycle', False)])

        add_tf(a.id, 'Sarah drinks coffee for breakfast.',
               'Select True or False based on the passage.', 1, 3, False)

        add_fill(a.id, 'Sarah practices playing the _______ for one hour.',
                 'Type the missing word from the passage.', 2, 4, 'guitar')

        add_mc(a.id, 'What does Sarah do in the evening?',
               'Choose the correct answer.', 2, 5,
               [('She watches TV', False),
                ('She helps her mother prepare dinner', True),
                ('She goes to the gym', False),
                ('She reads books', False)])

        add_short(a.id, 'Write 3-4 sentences describing YOUR daily routine using present simple tense.',
                  'Write in complete sentences. Use verbs like: wake up, eat, go, study, play.',
                  4, 6)

        db.session.commit()
        all_codes.append(('READ01BB', '3rd Secondary', 'Pre-Intermediate (B1)', a.title))
        print(f'  Created: {a.title}')

        # ---------------------------------------------------------------
        # 18. 3ro SECUNDARIA - B1: Relative Clauses and Connectors
        # ---------------------------------------------------------------
        a = Assessment(
            title='Relative Clauses and Linking Words',
            description='Practice using relative pronouns (who, which, that, where) and connectors (although, however, because, so).',
            assessment_type='exam',
            grade_id=get_grade(9).id,
            english_level_id=get_level('B1').id,
            teacher_id=tid,
            time_limit_minutes=30,
            show_results=True,
            access_code='RELA3SEC'
        )
        db.session.add(a)
        db.session.flush()

        add_mc(a.id, 'The woman _____ lives next door is a doctor.',
               'Choose the correct relative pronoun.', 2, 1,
               [('which', False), ('who', True), ('where', False), ('what', False)])

        add_mc(a.id, 'This is the book _____ I told you about.',
               'Choose the correct relative pronoun.', 2, 2,
               [('who', False), ('where', False), ('that', True), ('whom', False)])

        add_fill(a.id, 'Paris is the city ______ I was born.',
                 'Type the correct relative pronoun.', 2, 3, 'where')

        add_mc(a.id, 'I stayed home _____ I was feeling sick.',
               'Choose the correct connector.', 2, 4,
               [('although', False), ('because', True), ('however', False), ('but', False)])

        add_matching(a.id, 'Match each connector with its function.',
                     'Match correctly.', 4, 5,
                     [('Because', 'Gives a reason'), ('Although', 'Shows contrast'),
                      ('So', 'Shows result'), ('However', 'Shows contrast (formal)')])

        add_fill(a.id, 'He studied very hard; ______, he passed the exam.',
                 'Type the correct connector.', 2, 6, 'therefore')

        add_tf(a.id, '"Who" is used for people and "which" is used for things.',
               'Select True or False.', 1, 7, True)

        add_mc(a.id, '_____ it was raining, we went to the beach.',
               'Choose the correct connector.', 2, 8,
               [('Because', False), ('So', False), ('Although', True), ('Therefore', False)])

        add_mc(a.id, 'The restaurant _____ we had dinner was excellent.',
               'Choose the correct relative pronoun.', 2, 9,
               [('who', False), ('which', False), ('where', True), ('that', False)])

        add_short(a.id, 'Write 3 sentences using different relative pronouns (who, which, where).',
                  'Example: "My teacher, who is from England, speaks three languages." Write your own sentences.',
                  4, 10)

        db.session.commit()
        all_codes.append(('RELA3SEC', '3rd Secondary', 'Pre-Intermediate (B1)', a.title))
        print(f'  Created: {a.title}')

        # ---------------------------------------------------------------
        # 19. 4to SECUNDARIA - B1: Passive Voice and Modal Verbs
        # ---------------------------------------------------------------
        a = Assessment(
            title='Passive Voice and Modal Verbs',
            description='Practice transforming active sentences to passive and using modal verbs (must, should, might, can).',
            assessment_type='exam',
            grade_id=get_grade(10).id,
            english_level_id=get_level('B1').id,
            teacher_id=tid,
            time_limit_minutes=35,
            show_results=True,
            access_code='PASS4SEC'
        )
        db.session.add(a)
        db.session.flush()

        add_mc(a.id, 'The Mona Lisa _____ by Leonardo da Vinci.',
               'Choose the correct passive form.', 2, 1,
               [('painted', False), ('was painted', True),
                ('is painting', False), ('paints', False)])

        add_mc(a.id, 'English _____ in many countries around the world.',
               'Choose the correct passive form.', 2, 2,
               [('speaks', False), ('is spoken', True),
                ('speaking', False), ('spoke', False)])

        add_fill(a.id, 'The homework ______ (must / complete) before Friday.',
                 'Type the complete passive form with modal verb.', 2, 3,
                 'must be completed')

        add_mc(a.id, 'You _____ drive without a license. It\'s against the law.',
               'Choose the correct modal verb.', 2, 4,
               [('should', False), ('can', False), ("mustn't", True), ('might', False)])

        add_mc(a.id, 'You look tired. You _____ go to bed early tonight.',
               'Choose the correct modal verb for advice.', 2, 5,
               [('must', False), ('should', True), ('might', False), ('can', False)])

        add_matching(a.id, 'Match each modal verb with its meaning.',
                     'Match correctly.', 4, 6,
                     [('Must', 'Obligation / Necessity'), ('Should', 'Advice / Recommendation'),
                      ('Might', 'Possibility'), ('Can', 'Ability / Permission')])

        add_tf(a.id, 'In the passive voice, the object of the active sentence becomes the subject.',
               'Select True or False.', 1, 7, True)

        add_fill(a.id, 'Active: "They build houses." Passive: "Houses ______ built (by them)."',
                 'Type the correct passive form.', 2, 8, 'are')

        add_mc(a.id, 'The new hospital _____ next year.',
               'Choose the correct future passive.', 2, 9,
               [('will built', False), ('will be built', True),
                ('is building', False), ('built', False)])

        add_mc(a.id, 'She _____ be at home. I saw her car in the driveway.',
               'Choose the correct modal for deduction.', 2, 10,
               [('can\'t', False), ('might', True), ('shouldn\'t', False), ('mustn\'t', False)])

        add_short(a.id, 'Rewrite these sentences in passive voice:\n1. People speak Spanish in Peru.\n2. Alexander Graham Bell invented the telephone.\n3. Someone stole my bicycle.',
                  'Write the 3 sentences in passive form.', 6, 11)

        db.session.commit()
        all_codes.append(('PASS4SEC', '4th Secondary', 'Pre-Intermediate (B1)', a.title))
        print(f'  Created: {a.title}')

        # ---------------------------------------------------------------
        # 20. 4to SECUNDARIA - B2: Reading: Technology and Society
        # ---------------------------------------------------------------
        a = Assessment(
            title='Reading Comprehension: Technology and Society',
            description='Read an article about the impact of technology on modern society and answer comprehension questions.',
            assessment_type='exam',
            grade_id=get_grade(10).id,
            english_level_id=get_level('B2').id,
            teacher_id=tid,
            time_limit_minutes=40,
            show_results=True,
            access_code='TECH4SEC'
        )
        db.session.add(a)
        db.session.flush()

        article = (
            'Read the following article:\n\n'
            '"THE DIGITAL AGE: HOW TECHNOLOGY IS CHANGING OUR LIVES\n\n'
            'Over the past two decades, technology has transformed nearly every aspect of daily life. '
            'Smartphones, which were once considered luxury items, have become essential tools that '
            'most people carry everywhere. Social media platforms have changed the way we communicate, '
            'share information, and even form opinions.\n\n'
            'While technology has brought many benefits — such as instant access to information, '
            'improved healthcare, and more efficient communication — it has also raised serious '
            'concerns. Studies have shown that excessive screen time can lead to sleep problems, '
            'anxiety, and reduced physical activity, particularly among teenagers.\n\n'
            'Furthermore, the rise of artificial intelligence has sparked debates about the future '
            'of work. Many experts predict that AI will automate millions of jobs, while others '
            'argue that it will create new opportunities that we cannot yet imagine.\n\n'
            'Despite these challenges, most people agree that technology will continue to play an '
            'increasingly important role in our lives. The key is to find a balance between '
            'embracing innovation and protecting our well-being."'
        )

        add_mc(a.id, article + '\n\nAccording to the article, smartphones have become:',
               'Choose the correct answer.', 2, 1,
               [('Luxury items only for the rich', False),
                ('Essential tools for most people', True),
                ('Unnecessary devices', False),
                ('Dangerous for everyone', False)])

        add_mc(a.id, 'What negative effect of technology is mentioned for teenagers?',
               'Choose the correct answer.', 2, 2,
               [('Loss of money', False),
                ('Sleep problems, anxiety, and reduced physical activity', True),
                ('Poor grades at school', False),
                ('Loss of friends', False)])

        add_tf(a.id, 'All experts agree that AI will only destroy jobs without creating new ones.',
               'Select True or False based on the article.', 1, 3, False)

        add_mc(a.id, 'What does the author suggest is "the key"?',
               'Choose the correct answer.', 2, 4,
               [('Avoiding technology completely', False),
                ('Using only social media', False),
                ('Finding a balance between innovation and well-being', True),
                ('Working with AI', False)])

        add_fill(a.id, 'Social media platforms have changed the way we ______, share information, and form opinions.',
                 'Type the missing verb from the passage.', 2, 5, 'communicate')

        add_mc(a.id, 'The word "sparked" in the third paragraph is closest in meaning to:',
               'Choose the closest synonym.', 2, 6,
               [('Ended', False), ('Started / Initiated', True),
                ('Prevented', False), ('Ignored', False)])

        add_matching(a.id, 'Match the technology topic with its impact mentioned in the article.',
                     'Match each topic with its described effect.', 3, 7,
                     [('Smartphones', 'Became essential everyday tools'),
                      ('Social media', 'Changed how people communicate'),
                      ('AI', 'Sparked debates about the future of work')])

        add_short(a.id, 'Do you think technology has a positive or negative impact on teenagers? Write a paragraph (5-6 sentences) giving your opinion with examples.',
                  'Use linking words: In my opinion, Furthermore, However, For example, In conclusion.',
                  6, 8)

        db.session.commit()
        all_codes.append(('TECH4SEC', '4th Secondary', 'Intermediate (B2)', a.title))
        print(f'  Created: {a.title}')

        # ---------------------------------------------------------------
        # 21. 5to SECUNDARIA - B2: Conditionals and Reported Speech Exam
        # ---------------------------------------------------------------
        a = Assessment(
            title='Conditionals and Reported Speech Exam',
            description='Comprehensive exam covering first, second, and third conditionals, as well as reported speech.',
            assessment_type='exam',
            grade_id=get_grade(11).id,
            english_level_id=get_level('B2').id,
            teacher_id=tid,
            time_limit_minutes=45,
            shuffle_questions=True,
            show_results=True,
            access_code='ADVGRAM01'
        )
        db.session.add(a)
        db.session.flush()

        add_mc(a.id, 'If I _____ more money, I would travel around the world.',
               'Choose the correct verb form (Second Conditional).', 2, 1,
               [('have', False), ('had', True), ('will have', False), ('would have', False)])

        add_mc(a.id, 'If it rains tomorrow, we _____ the picnic.',
               'Choose the correct verb form (First Conditional).', 2, 2,
               [('cancel', False), ('will cancel', True), ('would cancel', False), ('cancelled', False)])

        add_mc(a.id, 'If she had studied harder, she _____ the exam.',
               'Choose the correct verb form (Third Conditional).', 2, 3,
               [('would pass', False), ('would have passed', True),
                ('will pass', False), ('passed', False)])

        add_fill(a.id, 'She said: "I am tired." -> She said that she _____ tired. (Reported Speech)',
                 'Type the correct reported form.', 2, 4, 'was')

        add_fill(a.id, 'He said: "I will call you tomorrow." -> He said that he _____ call me the next day.',
                 'Type the correct word.', 2, 5, 'would')

        add_tf(a.id, 'In the third conditional, we use "would have + past participle" in the result clause.',
               'Select True or False.', 1, 6, True)

        add_matching(a.id, 'Match each conditional type with its structure.',
                     'Match the left column with the right column.', 3, 7,
                     [('First Conditional', 'If + present simple, will + base verb'),
                      ('Second Conditional', 'If + past simple, would + base verb'),
                      ('Third Conditional', 'If + past perfect, would have + past participle')])

        add_short(a.id, 'Write one sentence using each type of conditional (First, Second, and Third). '
                  'Label each sentence.',
                  'Write 3 complete sentences, each clearly labeled.', 6, 8)

        db.session.commit()
        all_codes.append(('ADVGRAM01', '5th Secondary', 'Intermediate (B2)', a.title))
        print(f'  Created: {a.title}')

        # ---------------------------------------------------------------
        # 22. 5to SECUNDARIA - B2: Writing: Formal Email and Essay Structure
        # ---------------------------------------------------------------
        a = Assessment(
            title='Writing: Formal Email and Essay Structure',
            description='Practice writing formal emails and understanding essay organization with introduction, body, and conclusion.',
            assessment_type='exam',
            grade_id=get_grade(11).id,
            english_level_id=get_level('B2').id,
            teacher_id=tid,
            time_limit_minutes=45,
            show_results=True,
            access_code='WRIT5SEC'
        )
        db.session.add(a)
        db.session.flush()

        add_mc(a.id, 'Which greeting is appropriate for a formal email?',
               'Choose the most formal option.', 2, 1,
               [('Hey there!', False), ('Dear Mr. Johnson,', True),
                ('Hi buddy,', False), ('What\'s up,', False)])

        add_mc(a.id, 'Which closing is appropriate for a formal email?',
               'Choose the most formal option.', 2, 2,
               [('See ya!', False), ('Later!', False),
                ('Sincerely,', True), ('Bye bye!', False)])

        add_matching(a.id, 'Match each part of an essay with its description.',
                     'Match correctly.', 3, 3,
                     [('Introduction', 'Presents the topic and thesis statement'),
                      ('Body paragraphs', 'Develops arguments with evidence and examples'),
                      ('Conclusion', 'Summarizes main points and restates the thesis')])

        add_tf(a.id, 'In a formal email, you should use contractions like "don\'t" and "can\'t".',
               'Select True or False.', 1, 4, False)

        add_mc(a.id, 'Which sentence is the best thesis statement?',
               'Choose the most effective thesis.', 2, 5,
               [('This essay is about pollution.', False),
                ('Pollution is bad.', False),
                ('Industrial pollution is the leading cause of climate change and requires immediate government action.', True),
                ('I think pollution is a problem.', False)])

        add_mc(a.id, '"I am writing to inquire about..." is an example of:',
               'Choose the correct register.', 2, 6,
               [('Informal language', False), ('Slang', False),
                ('Formal language', True), ('Academic language', False)])

        add_fill(a.id, 'A topic ______ is the first sentence of a body paragraph that introduces the main idea.',
                 'Type the missing word.', 2, 7, 'sentence')

        add_ordering(a.id, 'Put the parts of a formal email in order.',
                     'Arrange from first to last.', 4, 8,
                     ['Subject line', 'Greeting (Dear...)', 'Body of the email', 'Closing (Sincerely,)'])

        add_short(a.id, 'Write a formal email to your school principal requesting permission to organize a cultural event. Include: greeting, reason for writing, details of the event, and a polite closing.',
                  'Use formal language. Start with "Dear Principal..." and end with "Sincerely, [Your name]." Write at least 6 sentences.',
                  8, 9)

        add_short(a.id, 'Write an introduction paragraph (4-5 sentences) for an essay about the following topic: "Should students wear uniforms at school?"',
                  'Include a hook, background information, and a clear thesis statement.',
                  6, 10)

        db.session.commit()
        all_codes.append(('WRIT5SEC', '5th Secondary', 'Intermediate (B2)', a.title))
        print(f'  Created: {a.title}')

        # ---------------------------------------------------------------
        # 23. 5to SECUNDARIA - C1: Advanced Grammar and Vocabulary
        # ---------------------------------------------------------------
        a = Assessment(
            title='Advanced Grammar: Mixed Tenses and Phrasal Verbs',
            description='Challenge yourself with advanced grammar including mixed tenses, phrasal verbs, and complex sentence structures.',
            assessment_type='exam',
            grade_id=get_grade(11).id,
            english_level_id=get_level('C1').id,
            teacher_id=tid,
            time_limit_minutes=50,
            shuffle_questions=True,
            show_results=True,
            access_code='ADVN5SEC'
        )
        db.session.add(a)
        db.session.flush()

        add_mc(a.id, 'By the time we arrived, the movie _____.',
               'Choose the correct tense.', 2, 1,
               [('already started', False), ('has already started', False),
                ('had already started', True), ('was already starting', False)])

        add_mc(a.id, 'I wish I _____ more time to study for the exam.',
               'Choose the correct form.', 2, 2,
               [('have', False), ('had', True), ('would have', False), ('will have', False)])

        add_fill(a.id, 'She has been working here ______ 2019.',
                 'Type "since" or "for".', 2, 3, 'since')

        add_mc(a.id, 'He finally _____ up smoking after trying for years.',
               'Choose the correct phrasal verb.', 2, 4,
               [('gave', True), ('put', False), ('took', False), ('made', False)])

        add_matching(a.id, 'Match each phrasal verb with its meaning.',
                     'Match correctly.', 4, 5,
                     [('Look up', 'Search for information'), ('Give up', 'Stop doing something'),
                      ('Put off', 'Postpone'), ('Turn down', 'Reject / Refuse')])

        add_fill(a.id, 'Not only ______ she speak English, but she also speaks French and Japanese.',
                 'Type the correct auxiliary for the inverted structure.', 2, 6, 'does')

        add_mc(a.id, 'Had I known about the traffic, I _____ left earlier.',
               'Choose the correct form.', 2, 7,
               [('will have', False), ('would', False),
                ('would have', True), ('had', False)])

        add_tf(a.id, '"Used to" describes a past habit or state that is no longer true.',
               'Select True or False.', 1, 8, True)

        add_mc(a.id, 'The manager insisted that every employee _____ on time.',
               'Choose the correct subjunctive form.', 2, 9,
               [('is', False), ('be', True), ('was', False), ('were', False)])

        add_mc(a.id, 'She\'s been feeling under the _____. She should see a doctor.',
               'Choose the correct idiom completion.', 2, 10,
               [('sky', False), ('weather', True), ('table', False), ('moon', False)])

        add_matching(a.id, 'Match each idiom with its meaning.',
                     'Match correctly.', 3, 11,
                     [('Break the ice', 'Start a conversation in a social setting'),
                      ('Hit the nail on the head', 'Say something exactly right'),
                      ('Piece of cake', 'Something very easy')])

        add_short(a.id, 'Write a paragraph (6-8 sentences) about a significant experience that changed your perspective on life. Use at least 3 different tenses and 2 phrasal verbs.',
                  'Demonstrate use of: past simple, past perfect, present perfect, at least 2 phrasal verbs. Underline or highlight the phrasal verbs.',
                  8, 12)

        db.session.commit()
        all_codes.append(('ADVN5SEC', '5th Secondary', 'Upper-Intermediate (C1)', a.title))
        print(f'  Created: {a.title}')

        # ==============================================================
        # SUMMARY
        # ==============================================================
        print('\n' + '=' * 65)
        print('SEED COMPLETE - All exam models created successfully!')
        print('=' * 65)
        print(f'\nTotal assessments created: {len(all_codes)}')
        print(f'\n{"CODE":<12} {"GRADE":<18} {"LEVEL":<28} TITLE')
        print('-' * 95)
        for code, grade, level, title in all_codes:
            print(f'{code:<12} {grade:<18} {level:<28} {title}')
        print('\nDefault admin login: admin / admin123')
        print('Access assessments at: http://localhost:5000')


if __name__ == '__main__':
    seed_assessments()
