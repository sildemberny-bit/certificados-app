from flask import Flask, render_template, request, redirect, url_for, session, send_file
import os
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from PIL import Image
import zipfile
from io import BytesIO

app = Flask(__name__)
app.secret_key = "emitte_2025_super_seguro"

usuarios = {"admin": "123"}
contador_certificados = 0


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form.get("email")
        password = request.form.get("password")

        if email in usuarios and usuarios[email] == password:
            session["usuario"] = email
            return redirect(url_for("dashboard"))

    return render_template("login.html")


@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        email = request.form.get("email")
        password = request.form.get("password")

        if email not in usuarios:
            usuarios[email] = password
            session["usuario"] = email
            return redirect(url_for("dashboard"))

    return render_template("register.html")


@app.route("/logout")
def logout():
    session.pop("usuario", None)
    return redirect(url_for("login"))


@app.route("/")
def dashboard():
    if "usuario" not in session:
        return redirect(url_for("login"))

    return render_template("dashboard.html", contador=contador_certificados)


@app.route("/certificados", methods=["GET", "POST"])
def certificados():
    global contador_certificados

    if "usuario" not in session:
        return redirect(url_for("login"))

    if request.method == "POST":

        fundo = request.files["fundo"]
        planilha = request.files["planilha"]
        texto = request.form.get("texto")

        df = pd.read_excel(planilha)

        zip_buffer = BytesIO()
        zip_file = zipfile.ZipFile(zip_buffer, "w")

        for index, row in df.iterrows():

            nome = str(row[0])

            pdf_buffer = BytesIO()
            c = canvas.Canvas(pdf_buffer, pagesize=A4)

            largura, altura = A4

            # Fundo
            img = Image.open(fundo)
            img.save("temp_fundo.png")
            c.drawImage("temp_fundo.png", 0, 0, largura, altura)

            # Texto centralizado
            c.setFont("Helvetica", 20)
            texto_final = texto.replace("{nome}", nome)
            largura_texto = c.stringWidth(texto_final, "Helvetica", 20)
            x = (largura - largura_texto) / 2
            y = altura / 2

            c.drawString(x, y, texto_final)

            c.save()

            pdf_buffer.seek(0)
            zip_file.writestr(f"{nome}.pdf", pdf_buffer.read())

            contador_certificados += 1

        zip_file.close()
        zip_buffer.seek(0)

        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name="certificados_emitte.zip",
            mimetype="application/zip"
        )

    return render_template("certificados.html")


if __name__ == "__main__":
    app.run(debug=True)
