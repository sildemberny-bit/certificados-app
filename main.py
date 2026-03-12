from flask import Flask, render_template, request, redirect, send_file, session
import pandas as pd
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import zipfile
import io

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

        imagem = Image.open(fundo)
        imagem = imagem.convert("RGB")

        buffer_zip = io.BytesIO()

        with zipfile.ZipFile(buffer_zip,"w") as zipf:

            for i, linha in df.iterrows():

                texto_certificado = texto

                for coluna in df.columns:

                    valor = str(linha[coluna])

                    texto_certificado = texto_certificado.replace(
                        "{" + coluna.upper() + "}",
                        f"<b>{valor}</b>"
                    )

                texto_certificado = texto_certificado.replace("\n","<br/><br/>")

                buffer_pdf = io.BytesIO()

                largura_pagina, altura_pagina = landscape(A4)

                c = canvas.Canvas(
                    buffer_pdf,
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

                p = Paragraph(texto_certificado, style)

                # calcular tamanho real do bloco
                w, h = p.wrap(largura_texto, altura_pagina)

                # área segura (evita invadir logos)
                topo_seguro = altura_pagina * 0.75

                if posicao_vertical == "superior":
                    y = topo_seguro - h

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

                buffer_pdf.seek(0)

                zipf.writestr(
                    f"certificado_{i+1}.pdf",
                    buffer_pdf.read()
                )

        buffer_zip.seek(0)

        return send_file(
            buffer_zip,
            as_attachment=True,
            download_name="certificados.zip",
            mimetype="application/zip"
        )

    return render_template("certificados.html")


if __name__ == "__main__":
    app.run(debug=True)
