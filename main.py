from flask import Flask, render_template, request, redirect, send_file, session
import pandas as pd
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import zipfile
import io
import os
import unicodedata
import re
import tempfile
import shutil

app = Flask(__name__)
app.secret_key = "emitte_secret"

USUARIO = "admin"
SENHA = "123"


@app.route("/")
def landing():
    return render_template("landing.html")


@app.route("/login", methods=["GET","POST"])
def login():

    if request.method == "POST":

        usuario = request.form.get("email")
        senha = request.form.get("password")

        if usuario == USUARIO and senha == SENHA:
            session["user"] = usuario
            return redirect("/certificados")

    return render_template("login.html")


@app.route("/logout")
def logout():

    session.pop("user",None)
    return redirect("/")


def limpar_nome_arquivo(nome):

    nome = nome.lower()

    nome = unicodedata.normalize("NFD", nome)
    nome = nome.encode("ascii", "ignore").decode("utf-8")

    nome = re.sub(r"[^a-z0-9 ]", "", nome)

    nome = nome.replace(" ", "_")

    return nome


def gerar_pdf(fundo, texto, fonte, alinhamento, posicao_vertical, caminho_pdf):

    imagem = Image.open(fundo)
    imagem = imagem.convert("RGB")

    largura_pagina, altura_pagina = landscape(A4)

    c = canvas.Canvas(
        caminho_pdf,
        pagesize=(largura_pagina, altura_pagina)
    )

    fundo_reader = ImageReader(imagem)

    c.drawImage(
        fundo_reader,
        0,
        0,
        width=largura_pagina,
        height=altura_pagina
    )

    largura_texto = largura_pagina * 0.85

    if alinhamento == "centro":
        alinh = TA_CENTER
    elif alinhamento == "esquerda":
        alinh = TA_LEFT
    else:
        alinh = TA_RIGHT

    style = ParagraphStyle(
        name="Certificado",
        fontName="Helvetica",
        fontSize=fonte + 6,
        leading=(fonte + 6) * 1.3,
        alignment=alinh
    )

    texto = texto.replace("\n","<br/>")

    p = Paragraph(texto, style)

    w, h = p.wrap(largura_texto, altura_pagina)

    if posicao_vertical == "superior":
        y = altura_pagina * 0.75
    elif posicao_vertical == "centro":
        y = (altura_pagina / 2) - (h / 2)
    else:
        y = altura_pagina * 0.30

    p.drawOn(
        c,
        (largura_pagina - largura_texto) / 2,
        y
    )

    c.save()


@app.route("/preview", methods=["POST"])
def preview():

    fundo = request.files["fundo"]
    texto = request.form["texto"]
    fonte = int(request.form["fonte"])
    alinhamento = request.form["alinhamento"]
    posicao_vertical = request.form["posicao_vertical"]

    buffer_pdf = io.BytesIO()

    imagem = Image.open(fundo)
    imagem = imagem.convert("RGB")

    largura_pagina, altura_pagina = landscape(A4)

    c = canvas.Canvas(buffer_pdf, pagesize=(largura_pagina, altura_pagina))

    fundo_reader = ImageReader(imagem)

    c.drawImage(
        fundo_reader,
        0,
        0,
        width=largura_pagina,
        height=altura_pagina
    )

    style = ParagraphStyle(
        name="Preview",
        fontName="Helvetica",
        fontSize=fonte + 6,
        leading=(fonte + 6) * 1.3,
        alignment=TA_CENTER
    )

    texto = texto.replace("\n","<br/>")

    p = Paragraph(texto, style)

    largura_texto = largura_pagina * 0.85

    w, h = p.wrap(largura_texto, altura_pagina)

    y = (altura_pagina / 2) - (h / 2)

    p.drawOn(
        c,
        (largura_pagina - largura_texto) / 2,
        y
    )

    c.save()

    buffer_pdf.seek(0)

    return send_file(buffer_pdf, mimetype="application/pdf")


@app.route("/certificados", methods=["GET","POST"])
def certificados():

    if "user" not in session:
        return redirect("/login")

    if request.method == "POST":

        fundo = request.files["fundo"]
        planilha = request.files["planilha"]

        texto = request.form["texto"]
        fonte = int(request.form["fonte"])
        alinhamento = request.form["alinhamento"]
        posicao_vertical = request.form["posicao_vertical"]

        df = pd.read_excel(planilha)

        pasta_temp = tempfile.mkdtemp()

        lista_pdfs = []

        for i, linha in df.iterrows():

            texto_certificado = texto

            for coluna in df.columns:

                valor = str(linha[coluna])

                texto_certificado = texto_certificado.replace(
                    "{" + coluna.upper() + "}",
                    f"<b>{valor}</b>"
                )

            nome_base = str(linha[df.columns[0]])
            nome_arquivo = limpar_nome_arquivo(nome_base) + ".pdf"

            caminho_pdf = os.path.join(pasta_temp, nome_arquivo)

            gerar_pdf(
                fundo,
                texto_certificado,
                fonte,
                alinhamento,
                posicao_vertical,
                caminho_pdf
            )

            lista_pdfs.append(caminho_pdf)

        caminho_zip = os.path.join(pasta_temp, "certificados.zip")

        with zipfile.ZipFile(caminho_zip, "w") as zipf:

            for pdf in lista_pdfs:
                zipf.write(pdf, os.path.basename(pdf))

        return send_file(
            caminho_zip,
            as_attachment=True,
            download_name="certificados.zip",
            mimetype="application/zip"
        )

    return render_template("certificados.html")


if __name__ == "__main__":

    port = int(os.environ.get("PORT", 10000))

    app.run(
        host="0.0.0.0",
        port=port
    )
