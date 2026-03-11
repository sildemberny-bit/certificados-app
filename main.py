from flask import Flask, render_template, request, redirect, url_for, session

app = Flask(__name__)
app.secret_key = "emitte_secret"

# LOGIN PADRÃO
USUARIO = "admin"
SENHA = "123"


@app.route("/", methods=["GET","POST"])
def login():

    if request.method == "POST":

        usuario = request.form.get("usuario")
        senha = request.form.get("senha")

        if usuario == USUARIO and senha == SENHA:
            session["logado"] = True
            return redirect("/certificados")

    return render_template("login.html")


@app.route("/certificados", methods=["GET","POST"])
def certificados():

    if not session.get("logado"):
        return redirect("/")

    return render_template("certificados.html")


@app.route("/logout")
def logout():

    session.clear()
    return redirect("/")


if __name__ == "__main__":
    app.run()
