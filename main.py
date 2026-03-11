from flask import Flask, render_template, request, redirect, session, send_file
import pandas as pd
import os
import gc

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape

app = Flask(__name__)
app.secret_key = "emitte_secret"

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "certificados"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# LOGIN
@app.route("/", methods=["GET","POST"])
def login():

    if request.method == "POST":

        email = request.form.get("email")
        password = request.form.get("password")

        if email == "admin" and password == "123":
            session["logado"] = True
            return redirect("/certificados")

    return render_template("login.html")


# LOGOUT
@app.route("/logout")
def logout():

    session.clear()
    return redirect("/")


# GUIA
@app.route("/guia")
def guia():

    return render_template("guia.html")


# PAGINA CERTIFICADOS
@app.route("/certificados", methods=["GET","POST"])
def certificados():

    if not session.get("logado"):
        return redirect("/")

    if request.method == "POST":

        excel = request.files["excel"]
        fundo = request.files["fundo"]
        texto_modelo = request.form["texto"]

        caminho_excel = os.path.join(UPLOAD_FOLDER, excel.filename)
        caminho_fundo = os.path.join(UPLOAD_FOLDER, fundo.filename)

        excel.save(caminho_excel)
        fundo.save(caminho_fundo)

        df = pd.read_excel(caminho_excel)

        contador = 0

        for index, row in df.iterrows():

            nome = str(row["NOME"])

            curso = str(row["CURSO"]) if "CURSO" in row else ""
            carga = str(row["CARGA"]) if "CARGA" in row else ""

            texto = texto_modelo.replace("{NOME}", nome)
            texto = texto.replace("{CURSO}", curso)
            texto = texto.replace("{CARGA}", carga)

            nome_arquivo = nome.replace(" ","_") + ".pdf"

            caminho_pdf = os.path.join(OUTPUT_FOLDER, nome_arquivo)

            c = canvas.Canvas(caminho_pdf, pagesize=landscape(A4))

            largura, altura = landscape(A4)

            c.drawImage(caminho_fundo, 0, 0, width=largura, height=altura)

            c.setFont("Helvetica", 20)

            c.drawCentredString(largura/2, altura/2, texto)

            c.save()

            # liberar memoria
            del c
            gc.collect()

            contador += 1

        print("NOVA GERAÇÃO REALIZADA:", contador)

        return f"{contador} certificados gerados com sucesso!"

    return render_template("certificados.html")


if __name__ == "__main__":
    app.run()
