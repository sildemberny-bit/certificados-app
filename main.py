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
import os
import unicodedata
import re
import tempfile
import datetime
import shutil

app = Flask(__name__)
app.secret_key = "emitte_secret"

USUARIO = "admin"
SENHA = "123"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PASTA_DOWNLOAD = os.path.join(BASE_DIR, "downloads")

os.makedirs(PASTA_DOWNLOAD, exist_ok=True)


@app.route("/")
def landing():
    return render_template("landing.html")


@app.route("/guia")
def guia():
    return render_template("guia.html")


@app.route("/login", methods=["GET", "POST"])
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
    session.pop("user", None)
    return redirect("/login")


def limpar_nome_arquivo(nome):

    nome = nome.lower()

    nome = unicodedata.normalize("NFD", nome)
    nome = nome.encode("ascii", "ignore").decode("utf-8")

    nome = re.sub(r"[^a-z0-9 ]", "", nome)

    nome = nome.replace(" ", "_")

    return nome


def substituir_campos(texto, linha):

    resultado = texto

    for coluna in linha.index:

        valor = str(linha[coluna])

        resultado = resultado.replace(
            "{" + coluna + "}",
            f"<b>{valor}</b>"
        )

        resultado = resultado.replace(
            "{" + coluna.lower() + "}",
            f"<b>{valor}</b>"
        )

        resultado = resultado.replace(
            "{" + coluna.upper() + "}",
            f"<b>{valor}</b>"
        )

    return resultado


def criar_estilo(fonte, alinhamento):

    if alinhamento == "centro":
        alinh = TA_CENTER

    elif alinhamento == "esquerda":
        alinh = TA_LEFT

    else:
        alinh = TA_RIGHT

    return ParagraphStyle(
        name="Certificado",
        fontName="Helvetica",
        fontSize=fonte,
        leading=fonte * 1.4,
        alignment=alinh
    )


def gerar_pdf_individual(
    imagem,
    linha,
    texto,
    fonte,
    alinhamento,
    posicao_vertical,
    ajuste_vertical,
    largura_texto_percent,
    caminho_pdf
):

    largura_pagina, altura_pagina = landscape(A4)

    c = canvas.Canvas(
        caminho_pdf,
        pagesize=(largura_pagina, altura_pagina)
    )

    fundo_reader = ImageReader(imagem)

    largura_texto = (
        largura_pagina * (largura_texto_percent / 100)
    )

    style = criar_estilo(
        fonte,
        alinhamento
    )

    texto_certificado = substituir_campos(
        texto,
        linha
    )

    texto_certificado = texto_certificado.replace(
        "\n",
        "<br/>"
    )

    c.drawImage(
        fundo_reader,
        0,
        0,
        width=largura_pagina,
        height=altura_pagina
    )

    p = Paragraph(
        texto_certificado,
        style
    )

    w, h = p.wrap(
        largura_texto,
        altura_pagina
    )

    if posicao_vertical == "superior":

        y = altura_pagina * 0.75

    elif posicao_vertical == "centro":

        y = (altura_pagina / 2) - (h / 2)

    else:

        y = altura_pagina * 0.30

    y = y + ajuste_vertical

    p.drawOn(
        c,
        (largura_pagina - largura_texto) / 2,
        y
    )

    c.showPage()
    c.save()


def detectar_coluna_nome(df):

    for col in df.columns:

        col_norm = unicodedata.normalize(
            "NFD",
            col
        )

        col_norm = col_norm.encode(
            "ascii",
            "ignore"
        ).decode("utf-8")

        col_norm = col_norm.lower()

        if "nome" in col_norm:
            return col

    return df.columns[0]


@app.route("/download/<arquivo>")
def baixar(arquivo):

    caminho = os.path.join(
        PASTA_DOWNLOAD,
        arquivo
    )

    return send_file(
        caminho,
        as_attachment=True
    )


@app.route("/certificados", methods=["GET", "POST"])
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

        ajuste_vertical = int(
            request.form["ajuste_vertical"]
        )

        largura_texto_percent = int(
            request.form["largura_texto"]
        )

        df = pd.read_excel(planilha)

        imagem = Image.open(fundo)
        imagem = imagem.convert("RGB")

        pasta_temp = tempfile.mkdtemp()

        quantidade = len(df)

        data = datetime.date.today().strftime(
            "%Y-%m-%d"
        )

        nome_zip = (
            f"certificados_emitte_{quantidade}_{data}.zip"
        )

        caminho_zip = os.path.join(
            PASTA_DOWNLOAD,
            nome_zip
        )

        coluna_nome = detectar_coluna_nome(df)

        try:

            # O ZIP não comprime os PDFs.
            # O ponto principal aqui é usar zipf.write(),
            # evitando carregar cada PDF inteiro na memória.

            with zipfile.ZipFile(
                caminho_zip,
                "w",
                compression=zipfile.ZIP_STORED
            ) as zipf:

                for i, linha in df.iterrows():

                    nome = str(
                        linha[coluna_nome]
                    )

                    nome_limpo = limpar_nome_arquivo(
                        nome
                    )

                    caminho_pdf = os.path.join(
                        pasta_temp,
                        f"{nome_limpo}.pdf"
                    )

                    gerar_pdf_individual(
                        imagem,
                        linha,
                        texto,
                        fonte,
                        alinhamento,
                        posicao_vertical,
                        ajuste_vertical,
                        largura_texto_percent,
                        caminho_pdf
                    )

                    # Adiciona o PDF diretamente ao ZIP.
                    # Não usamos f.read(), evitando carregar
                    # o arquivo inteiro novamente na memória.

                    zipf.write(
                        caminho_pdf,
                        arcname=os.path.basename(caminho_pdf)
                    )

                    # O PDF temporário não precisa permanecer
                    # depois de ser adicionado ao ZIP.

                    os.remove(caminho_pdf)

        finally:

            # Remove a pasta temporária ao terminar.
            shutil.rmtree(
                pasta_temp,
                ignore_errors=True
            )

        return render_template(
            "download.html",
            arquivo=nome_zip
        )

    return render_template(
        "certificados.html"
    )


if __name__ == "__main__":

    port = int(
        os.environ.get(
            "PORT",
            10000
        )
    )

    app.run(
        host="0.0.0.0",
        port=port
    )
