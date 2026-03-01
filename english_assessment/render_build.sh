#!/usr/bin/env bash
# Render build script
set -o errexit

pip install -r requirements.txt

# Seed database with sample assessments on first deploy
python -c "
from app import app
from models import db, Grade
with app.app_context():
    if not Grade.query.first():
        from seed_data import seed_assessments
        seed_assessments()
        print('Database seeded successfully.')
    else:
        print('Database already seeded.')
"
