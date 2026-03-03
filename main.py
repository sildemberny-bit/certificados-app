from flask import Flask, render_template, request, redirect, url_for, session
import os

app = Flask(__name__)
app.secret_key = "emitte_super_secreto"

# Contador simples em memória
total_gerado = 0
ultima_geracao = "Nenhuma ainda"


# =========================
# LOGIN
# =========================
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = request.form.get("usuario")
        senha = request.form.get("senha")

        if usuario == "admin" and senha == "123":
            session["usuario"] = usuario
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


# =========================
# RODAR APP
# =========================
if __name__ == "__main__":
    app.run(debug=True)
