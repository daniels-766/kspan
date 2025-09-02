from extensions import db
from models import User
from werkzeug.security import generate_password_hash
from app import app

username = "admin"
email = "admin@gmail.com"
phone = "081234567890"
password = "admin123"

with app.app_context():
    if User.query.filter((User.username == username) | (User.email == email)).first():
        print("Username atau email sudah terdaftar.")
    else:
        hashed_pw = generate_password_hash(password)
        new_admin = User(
            username=username,
            email=email,
            phone=phone,
            password=hashed_pw,
            role="admin"
        )
        db.session.add(new_admin)
        db.session.commit()
        print(f"Admin '{username}' berhasil dibuat!")
