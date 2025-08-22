# manage.py
from getpass import getpass
from werkzeug.security import generate_password_hash
from app import app
from models import db, User

with app.app_context():
    email = input("Email: ").strip().lower()
    username = input("Felhasználónév: ").strip()
    pw = getpass("Jelszó: ")
    u = User(email=email, username=username, password_hash=generate_password_hash(pw))
    db.session.add(u)
    db.session.commit()
    print("Felhasználó létrehozva.")
