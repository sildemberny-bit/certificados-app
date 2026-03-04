from flask import Flask, render_template, request, redirect, url_for, session

app = Flask(__name__)
app.secret_key = "emitte_2025_super_seguro"


# =========================
# LOGIN
# =========================
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form.get("email")
        password = request.form.get("password")

        if email == "admin" and password == "123":
            session["usuario"] = email
            return redirect(url_for("dashboard"))

    return render_template("login.html")


# =========================
# LOGOUT
# =========================
@app.route("/logout")
def logout():
    session.pop("usuario", None)
    return redirect(url_for("login"))


# =========================
# DASHBOARD
# =========================
@app.route("/")
def dashboard():
    if "usuario" not in session:
        return redirect(url_for("login"))

    return render_template("dashboard.html")


# =========================
# PÁGINA DE CERTIFICADOS
# =========================
@app.route("/certificados")
def certificados():
    if "usuario" not in session:
        return redirect(url_for("login"))

    return render_template("certificados.html")


if __name__ == "__main__":
    app.run(debug=True)
