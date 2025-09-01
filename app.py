import os
from io import BytesIO
from datetime import datetime

from flask import (
    Flask, render_template, request, redirect, url_for, flash,
    send_file, session, current_app, abort
)
from flask_login import (
    LoginManager, login_user, login_required, logout_user, current_user
)
from flask_mail import Mail, Message
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash

from dotenv import load_dotenv

from models import db, User, Upload, Result
from compute import process_excel_to_results, export_results_excel, export_results_pdfs
from admin import init_admin  # -> itt történik az Admin felület teljes bekötése

# ---- .env betöltése ----
load_dotenv()

mail = Mail()
login_manager = LoginManager()

# Engedélyezett kiterjesztések (feltöltéshez)
ALLOWED_EXT = {".xlsx", ".xls"}


# ----- App factory -----
def create_app():
    app = Flask(__name__)

    # Alap config
    app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-secret")
    app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get("DATABASE_URL", "sqlite:///app.db")
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
    app.config["UPLOAD_FOLDER"] = os.path.join(os.getcwd(), "uploads")
    app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20 MB
    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)

    # E-mail config (env-ből)
    app.config.update(
        MAIL_SERVER=os.environ.get("MAIL_SERVER", "smtp.gmail.com"),
        MAIL_PORT=int(os.environ.get("MAIL_PORT", "587")),
        MAIL_USE_TLS=os.environ.get("MAIL_USE_TLS", "true").lower() == "true",
        MAIL_USE_SSL=os.environ.get("MAIL_USE_SSL", "false").lower() == "true",
        MAIL_USERNAME=os.environ.get("MAIL_USERNAME"),
        MAIL_PASSWORD=os.environ.get("MAIL_PASSWORD"),
        MAIL_DEFAULT_SENDER=os.environ.get("MAIL_DEFAULT_SENDER"),
        ADMIN_EMAIL=os.environ.get("ADMIN_EMAIL"),
    )

    # Bővítmények bekötése
    db.init_app(app)
    mail.init_app(app)
    login_manager.init_app(app)
    login_manager.login_view = "login"

    # Táblák létrehozása és Admin felület inicializálása
    with app.app_context():
        db.create_all()
        init_admin(app)  # <- CSAK innen, nincs másik Admin példány
        seed_admin_from_env()
    return app

def seed_admin_from_env():
    email = os.getenv("ADMIN_SEED_EMAIL")
    username = os.getenv("ADMIN_SEED_USERNAME")
    password = os.getenv("ADMIN_SEED_PASSWORD")
    if not (email and username and password):
        return
    # már létezik?
    if User.query.filter((User.email == email) | (User.username == username)).first():
        return

    u = User(
        email=email.strip().lower(),
        username=username.strip(),
        password_hash=generate_password_hash(password),
        is_admin=True,
    )
    db.session.add(u)
    db.session.commit()
    print(">>> Admin user seeded:", email)

app = create_app()


@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))


# ----- I18N (nagyon egyszerű) -----
STRINGS = {
    "hu": {
        "site_title": "Antropometriai mérések",
        "hero_title": "Sporttudományi mérések és visszajelzés",
        "hero_sub": "Tölts fel méréseket Excelben, és kapj részletes eredményeket.",
        "login": "Bejelentkezés",
        "register": "Regisztráció",
        "logout": "Kijelentkezés",
        "dashboard": "Vezérlőpult",
        "download_template": "Excel sablon letöltése",
        "upload_excel": "Excel feltöltése",
        "your_uploads": "Feltöltéseid",
        "results_xlsx": "Eredmények (Excel)",
        "results_pdfs": "Egyéni visszajelzőlapok (PDF)",
        "lang_hu": "Magyar",
        "lang_en": "English",
        "request_access": "Hozzáférés igénylése",
        "about_title": "Antropometria kalkulátor",
        "about_p1": "Az oldal antropometriai és sporttudományi mérések feldolgozását és visszajelzését támogatja.",
        "about_p2": "Az adatok Excel sablonból tölthetők fel; a rendszer kiszámítja a fő mutatókat (BMI, endo/mezo/ekto, testzsír%, PHV, stb.).",
        "about_li1": "Iskolai/sportegyesületi mérések egységes feldolgozása",
        "about_li2": "Standardizált visszajelzőlapok (PDF)",
        "about_li3": "Időbeli fejlődés követése",
        "about_p3": "A hozzáférés engedélyezéshez kötött. Amennyiben szeretnéd használni, küldj igénylést. További információ: birozsolt.pszi@gmail.com",
        "features_title": "Funkciók",
        "f_template": "Excel sablon letöltése",
        "f_upload": "Kitöltött Excel feltöltése és feldolgozása",
        "f_results_excel": "Csoportos eredmények exportja (Excel)",
        "f_results_pdf": "Egyéni visszajelző lapok exportja (PDF)",
        "f_history": "Saját feltöltések visszanézése",
        "f_multilang": "Kétnyelvű felület (HU/EN)",
        "request_access_title": "Hozzáférés igénylése",
        "field_name": "Név",
        "field_email": "E-mail",
        "field_org": "Szervezet",
        "field_message": "Üzenet",
        "btn_send": "Küldés",
        "login_email": "E-mail",
        "login_password": "Jelszó",
        "login_button": "Bejelentkezés",
    },
    "en": {
        "site_title": "Anthropometry",
        "hero_title": "Sports Science Measurements & Feedback",
        "hero_sub": "Upload measurements in Excel and get detailed results.",
        "login": "Log in",
        "register": "Sign up",
        "logout": "Log out",
        "dashboard": "Dashboard",
        "download_template": "Download Excel template",
        "upload_excel": "Upload Excel",
        "your_uploads": "Your uploads",
        "results_xlsx": "Results (Excel)",
        "results_pdfs": "Individual feedback (PDFs)",
        "lang_hu": "Magyar",
        "lang_en": "English",
        "request_access": "Request access",
        "about_title": "Anthropometry Calculator",
        "about_p1": "This app processes anthropometric & sport science measurements and provides feedback.",
        "about_p2": "Data is uploaded via an Excel template; the system computes key indices (BMI, endo/meso/ecto, body fat %, PHV, etc.).",
        "about_li1": "Consistent processing for school/club assessments",
        "about_li2": "Standardized individual feedback (PDF)",
        "about_li3": "Track progress over time",
        "about_p3": "Access is permission-based. If you’d like to use it, please submit a request. For additional information: birozsolt.pszi@gmail.com",
        "features_title": "Features",
        "f_template": "Download Excel template",
        "f_upload": "Upload filled Excel and process",
        "f_results_excel": "Export group results (Excel)",
        "f_results_pdf": "Export individual feedback (PDF)",
        "f_history": "View your previous uploads",
        "f_multilang": "Bilingual interface (HU/EN)",
        "request_access_title": "Request access",
        "field_name": "Name",
        "field_email": "Email",
        "field_org": "Organization",
        "field_message": "Message",
        "btn_send": "Send",
        "login_email": "Email",
        "login_password": "Password",
        "login_button": "Log in",
    }
}


@app.context_processor
def inject_i18n():
    lang = session.get("lang", "hu")
    return dict(t=STRINGS.get(lang, STRINGS["hu"]), lang=lang)


@app.route("/lang/<code>")
def set_lang(code):
    if code in STRINGS:
        session["lang"] = code
    return redirect(request.referrer or url_for("index"))


# ----- Public pages -----
@app.route("/")
def index():
    return render_template("index.html")


# ----- Auth -----
@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        email = request.form["email"].strip().lower()
        username = request.form["username"].strip()
        password = request.form["password"]
        if User.query.filter((User.email == email) | (User.username == username)).first():
            flash("Már létezik ilyen felhasználó / User already exists.", "danger")
            return redirect(url_for("register"))
        user = User(email=email, username=username, password_hash=generate_password_hash(password))
        db.session.add(user)
        db.session.commit()
        flash("Sikeres regisztráció. Jelentkezz be! / Registration successful.", "success")
        return redirect(url_for("login"))
    return render_template("auth_register.html")


@app.route("/request-access", methods=["GET", "POST"])
def request_access():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        email = request.form.get("email", "").strip()
        org = request.form.get("org", "").strip()
        message_txt = request.form.get("message", "").strip()

        admin_rcpt = app.config.get("ADMIN_EMAIL")
        if not admin_rcpt:
            flash("Admin e-mail nincs beállítva (ADMIN_EMAIL).", "danger")
            return redirect(url_for("request_access"))

        try:
            msg = Message(
                subject="Új hozzáférés igénylés",
                recipients=[admin_rcpt],
                body=(
                    f"Név: {name}\n"
                    f"Email: {email}\n"
                    f"Szervezet: {org}\n"
                    f"Üzenet:\n{message_txt}\n"
                    f"Időpont: {datetime.utcnow().isoformat()}Z"
                ),
            )
            mail.send(msg)
            flash("Köszönjük! Hamarosan felvesszük veled a kapcsolatot.", "success")
            return redirect(url_for("index"))
        except Exception as e:
            flash(f"Nem sikerült elküldeni az igényt: {e}", "danger")
            return redirect(url_for("request_access"))

    return render_template("request_access.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form["email"].strip().lower()
        password = request.form["password"]
        user = User.query.filter_by(email=email).first()
        if not user or not check_password_hash(user.password_hash, password):
            flash("Hibás email/jelszó. / Invalid credentials.", "danger")
            return redirect(url_for("login"))
        login_user(user)
        return redirect(url_for("dashboard"))
    return render_template("auth_login.html")


@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("index"))


# ----- Dashboard -----
@app.route("/dashboard")
@login_required
def dashboard():
    uploads = (
        Upload.query
        .filter_by(user_id=current_user.id)
        .order_by(Upload.created_at.desc())
        .all()
    )
    return render_template("dashboard.html", uploads=uploads)


@app.route("/download/template")
@login_required
def download_template():
    lang = session.get("lang", "hu").lower()
    # opcionális felülírás: ?lang=en
    lang = request.args.get("lang", lang).lower()
    fname = "template_en.xlsx" if lang == "en" else "template_hu.xlsx"
    path = os.path.join(current_app.root_path, "static", fname)
    return send_file(path, as_attachment=True, download_name="template.xlsx")



@app.route("/upload", methods=["POST"])
@login_required
def upload_excel():
    file = request.files.get("file")
    if not file or not file.filename:
        flash("Nincs kiválasztott fájl. / No file selected.", "warning")
        return redirect(url_for("dashboard"))

    filename = secure_filename(file.filename)
    ext = os.path.splitext(filename)[1].lower()

    # 1) kiterjesztés ellenőrzés még mentés előtt
    if ext not in ALLOWED_EXT:
        flash("Csak .xlsx vagy .xls tölthető fel.", "warning")
        return redirect(url_for("dashboard"))

    # 2) mentés csak ezután
    saved_name = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{filename}"
    saved_path = os.path.join(app.config["UPLOAD_FOLDER"], saved_name)
    file.save(saved_path)

    # 3) feldolgozás
    try:
        rows = process_excel_to_results(saved_path, current_user.id)
    except Exception as e:
        # ha feldolgozás közben dől el, töröljük a feltöltött fájlt
        try:
            os.remove(saved_path)
        except OSError:
            pass
        flash(f"Feldolgozási hiba: {e}", "danger")
        return redirect(url_for("dashboard"))

    # 4) DB mentés: Upload + Results
    up = Upload(user_id=current_user.id, original_filename=filename, stored_path=saved_path)
    db.session.add(up)
    db.session.flush()  # upload_id

    for r in rows:
        r.upload_id = up.id
        db.session.add(r)

    db.session.commit()
    flash("Feltöltés és feldolgozás kész. / Upload processed.", "success")
    return redirect(url_for("dashboard"))


@app.route("/uploads/<int:upload_id>")
@login_required
def view_upload(upload_id):
    up = Upload.query.filter_by(id=upload_id, user_id=current_user.id).first_or_404()
    results = Result.query.filter_by(upload_id=up.id).all()
    return render_template("uploads.html", up=up, results=results)


@app.route("/uploads/<int:upload_id>/export/xlsx")
@login_required
def export_upload_xlsx(upload_id):
    up = Upload.query.filter_by(id=upload_id, user_id=current_user.id).first_or_404()
    results = Result.query.filter_by(upload_id=up.id).all()
    xls_bytes, fname = export_results_excel(results)
    return send_file(
        BytesIO(xls_bytes),
        as_attachment=True,
        download_name=fname,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/uploads/<int:upload_id>/export/pdfs")
@login_required
def export_upload_pdfs(upload_id):
    up = Upload.query.filter_by(id=upload_id, user_id=current_user.id).first_or_404()
    results = Result.query.filter_by(upload_id=up.id).all()

    # a felület aktuális nyelve
    lang = session.get("lang", "hu").lower()
    # opcionális felülírás query-ből: ?lang=en
    lang = request.args.get("lang", lang).lower()
    if lang not in ("hu", "en"):
        lang = "hu"

    zip_bytes, zip_name = export_results_pdfs(results, lang=lang)
    return send_file(
        BytesIO(zip_bytes),
        as_attachment=True,
        download_name=zip_name,
        mimetype="application/zip"
    )


@app.errorhandler(403)
def forbidden(e):
    return render_template("403.html"), 403


if __name__ == "__main__":
    app.run(debug=True)
