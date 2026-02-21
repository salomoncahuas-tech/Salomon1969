#!/usr/bin/env python3
"""
English Assessment Platform - Entry Point

Usage:
    python run.py              # Start the application
    python run.py --seed       # Start and seed sample data
    python run.py --port 8080  # Start on a custom port

Default admin credentials:
    Username: admin
    Password: admin123
"""

import argparse
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))


def main():
    parser = argparse.ArgumentParser(description='English Assessment Platform')
    parser.add_argument('--seed', action='store_true',
                        help='Seed the database with sample assessments')
    parser.add_argument('--port', type=int, default=5000,
                        help='Port to run the server on (default: 5000)')
    parser.add_argument('--host', default='0.0.0.0',
                        help='Host to bind to (default: 0.0.0.0)')
    args = parser.parse_args()

    from app import app

    if args.seed:
        from seed_data import seed_assessments
        seed_assessments()

    print(f'\n  English Assessment Platform')
    print(f'  Running on http://{args.host}:{args.port}')
    print(f'  Admin login: /admin/login (admin / admin123)\n')

    app.run(debug=True, host=args.host, port=args.port)


if __name__ == '__main__':
    main()
