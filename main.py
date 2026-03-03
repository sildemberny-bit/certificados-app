import os
import uuid
import zipfile
import sqlite3
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, send_file
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from PIL import Image, ImageDraw, ImageFont

app = Flask(__name__)
app.secret_key = "supersecretkey"

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "login"

DATABASE = "database.db"

def init_db():
    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()

    c.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        email TEXT UNIQUE,
        password TEXT,
        total_generated INTEGER DEFAULT 0
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS generations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        quantity INTEGER,
        date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    conn.commit()
    conn.close()

init_db()

class User(UserMixin):
    def __init__(self, id, name, email, password, total_generated):
        self.id = id
        self.name = name
        self.email = email
        self.password = password
        self.total_generated = total_generated

@login_manager.user_loader
def load_user(user_id):
    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    c.execute("SELECT * FROM users WHERE id = ?", (user_id,))
    user = c.fetchone()
    conn.close()
    if user:
        return User(*user)
    return None

@app.route("/register", methods=["GET","POST"])
def register():
    if request.method == "POST":
        name = request.form["name"]
        email = request.form["email"]
        password = generate_password_hash(request.form["password"])

        conn = sqlite3.connect(DATABASE)
        c = conn.cursor()
        c.execute("INSERT INTO users (name,email,password) VALUES (?,?,?)",
                  (name,email,password))
        conn.commit()
        conn.close()

        return redirect(url_for("login"))

    return render_template("register.html")

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        email = request.form["email"]
        password = request.form["password"]

        conn = sqlite3.connect(DATABASE)
        c = conn.cursor()
        c.execute("SELECT * FROM users WHERE email = ?", (email,))
        user = c.fetchone()
        conn.close()

        if user and check_password_hash(user[3], password):
            login_user(User(*user))
            return redirect(url_for("dashboard"))

    return render_template("login.html")

@app.route("/dashboard")
@login_required
def dashboard():
    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    c.execute("SELECT SUM(quantity) FROM generations WHERE user_id = ?", (current_user.id,))
    total = c.fetchone()[0] or 0

    c.execute("SELECT quantity,date FROM generations WHERE user_id = ? ORDER BY date DESC", (current_user.id,))
    history = c.fetchall()
    conn.close()

    return render_template("dashboard.html", total=total, history=history)

@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login"))

UPLOAD_FOLDER = "uploads"
GENERATED_FOLDER = "generated_certificates"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)

@app.route("/", methods=["GET","POST"])
@login_required
def index():
    if request.method == "POST":
        excel = request.files["excel"]
        background = request.files["background"]

        if excel and background:
            df = pd.read_excel(excel)
            quantity = len(df)

            conn = sqlite3.connect(DATABASE)
            c = conn.cursor()
            c.execute("INSERT INTO generations (user_id, quantity) VALUES (?,?)",
                      (current_user.id, quantity))
            conn.commit()
            conn.close()

            return redirect(url_for("dashboard"))

    return render_template("index.html")

if __name__ == "__main__":
    app.run()
