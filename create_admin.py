# create_admin.py
from getpass import getpass
from werkzeug.security import generate_password_hash
from app import create_app
from models import db, User

app = create_app()

with app.app_context():
    email = input("Admin email: ").strip().lower()
    username = input("Admin felhasználónév: ").strip()
    password = getpass("Admin jelszó: ")

    if User.query.filter((User.email == email) | (User.username == username)).first():
        print("HIBA: már létezik ilyen email vagy felhasználónév.")
    else:
        u = User(
            email=email,
            username=username,
            password_hash=generate_password_hash(password),
            is_admin=True,
        )
        db.session.add(u)
        db.session.commit()
        print("OK: admin felhasználó létrehozva.")
