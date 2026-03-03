from flask import Flask, render_template, request, redirect, url_for, session

app = Flask(__name__)
app.secret_key = "emitte_super_secreto"

total_gerado = 0
ultima_geracao = "Nenhuma ainda"

# =========================
# LOGIN
# =========================
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form.get("email")
        password = request.form.get("password")

        # LOGIN SIMPLES PARA TESTE
        if email == "admin@emitte.com" and password == "123":
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
@app.route("/", methods=["GET", "POST"])
def dashboard():
    global total_gerado, ultima_geracao

    if "usuario" not in session:
        return redirect(url_for("login"))

    if request.method == "POST":
        total_gerado += 1
        ultima_geracao = "Agora mesmo"

    return render_template(
        "index.html",
        total_gerado=total_gerado,
        ultima_geracao=ultima_geracao
    )


if __name__ == "__main__":
    app.run(debug=True)
