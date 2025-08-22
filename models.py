from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from datetime import datetime

db = SQLAlchemy()

class User(db.Model, UserMixin):
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(255), unique=True, nullable=False)
    username = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    is_admin = db.Column(db.Boolean, nullable=False, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    uploads = db.relationship("Upload", back_populates="user", cascade="all, delete-orphan")

    def __repr__(self):
        return f"<User {self.username}>"
class Upload(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    original_filename = db.Column(db.String(255))
    stored_path = db.Column(db.String(500))

    user = db.relationship("User", back_populates="uploads")

    def __repr__(self):
        return f"<Upload {self.id} by {self.user_id}>"

class Result(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    upload_id = db.Column(db.Integer, db.ForeignKey("upload.id"), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)

    # Bemeneti meta
    name = db.Column(db.String(255))
    birth_date = db.Column(db.Date)
    meas_date = db.Column(db.Date)
    ttm = db.Column(db.Float)  # cm
    tts = db.Column(db.Float)  # kg

    # Számított értékek
    ca_years = db.Column(db.Float)
    plx = db.Column(db.Float)
    mk_raw = db.Column(db.Float)
    mk = db.Column(db.Float)
    mk_corr_factor = db.Column(db.Float)
    vttm = db.Column(db.Float)
    sum6 = db.Column(db.Float)
    bodyfat_percent = db.Column(db.Float)
    bmi = db.Column(db.Float)
    bmi_cat = db.Column(db.String(120))
    endo = db.Column(db.Float)
    endo_cat = db.Column(db.String(120))
    mezo = db.Column(db.Float)
    mezo_cat = db.Column(db.String(120))
    ekto = db.Column(db.Float)
    ekto_cat = db.Column(db.String(120))
    phv = db.Column(db.Float)
    phv_cat = db.Column(db.String(120))

class AccessRequest(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    name = db.Column(db.String(255))
    email = db.Column(db.String(255))
    org = db.Column(db.String(255))
    message = db.Column(db.Text)
