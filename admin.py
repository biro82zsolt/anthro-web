# admin.py
from flask_admin import Admin
from flask_admin.contrib.sqla import ModelView
from wtforms import PasswordField
from wtforms.validators import Optional, Length, Email, DataRequired
from flask_login import current_user
from werkzeug.security import generate_password_hash
from markupsafe import Markup
from flask import url_for, flash, redirect

from models import db, User, Upload, Result


class SecureModelView(ModelView):
    """Csak bejelentkezett admin érheti el az admin nézeteket."""
    page_size = 25
    can_view_details = True

    def is_accessible(self):
        return current_user.is_authenticated and getattr(current_user, "is_admin", False)

    def inaccessible_callback(self, name, **kwargs):
        flash("Nincs jogosultság az admin felülethez.", "danger")
        return redirect(url_for("login"))


class UserAdmin(SecureModelView):
    column_list = ("id", "email", "username", "is_admin", "created_at")
    column_searchable_list = ("email", "username")
    column_filters = ("is_admin",)
    can_view_details = True
    details_modal = True

    # ne mutassuk a hash-t, és ne szerkesszük a relációt innen
    form_excluded_columns = ("password_hash", "created_at", "uploads")

    # űrlapmezők: plusz egy jelszómező
    form_columns = ("email", "username", "is_admin", "password")
    form_extra_fields = {
        "password": PasswordField(
            "New password (optional)",
            validators=[Optional(), Length(min=6, message="Min. 6 karakter")]
        )
    }
    form_args = {
        "email": {"label": "E-mail", "validators": [DataRequired(), Email()]},
        "username": {"label": "Felhasználónév", "validators": [DataRequired(), Length(min=3)]},
    }

    def on_model_change(self, form, model, is_created):
        # csak akkor hash-eljük, ha adott meg új jelszót
        if hasattr(form, "password") and form.password.data:
            model.password_hash = generate_password_hash(form.password.data)

    def delete_model(self, model):
        # ne tudd törölni saját magad
        if current_user.is_authenticated and model.id == current_user.id:
            flash("Saját fiókodat nem törölheted.", "warning")
            return False
        return super().delete_model(model)


class UploadAdmin(SecureModelView):
    column_list = ("id", "user", "original_filename", "stored_path", "created_at")
    column_searchable_list = ("original_filename",)
    column_filters = ("created_at", "user.username")
    can_create = False
    can_edit = False
    can_view_details = True
    details_modal = True

    # gyors link az adott feltöltés oldalára
    def _link_to_upload(self, ctx, model, name):
        # /uploads/<id>
        link = url_for("view_upload", upload_id=model.id)
        return Markup(f'<a href="{link}" class="btn btn-sm btn-primary">Open</a>')

    column_formatters = {
        "stored_path": lambda v, c, m, p: Markup(f"<code>{m.stored_path}</code>"),
        "id": _link_to_upload,  # az ID oszlopban gombot mutatunk
    }


class ResultAdmin(SecureModelView):
    column_list = ("id", "upload_id")
    can_create = False
    can_edit = False
    can_view_details = True
    details_modal = True


def init_admin(app):
    admin = Admin(app, name="Admin", template_mode="bootstrap4")
    admin.add_view(UserAdmin(User, db.session, category="Manage"))
    admin.add_view(UploadAdmin(Upload, db.session, category="Manage"))
    admin.add_view(ResultAdmin(Result, db.session, category="Manage"))
    return admin
