from flask import Flask, render_template, request, send_file
import pandas as pd
import os
import uuid
import zipfile
from PIL import Image, ImageDraw, ImageFont

app = Flask(__name__)

PASTA_TEMP = "temp"

if not os.path.exists(PASTA_TEMP):
    os.makedirs(PASTA_TEMP)


def encontrar_coluna_nome(df):

    for col in df.columns:
        if str(col).strip().lower() == "nome":
            return col

    raise Exception("Coluna NOME não encontrada na planilha.")


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/login")
def login():
    return render_template("login.html")


@app.route("/guia")
def guia():
    return render_template("guia.html")


@app.route("/certificados")
def certificados():
    return render_template("certificados.html")


@app.route("/gerar", methods=["POST"])
def gerar():

    fundo = request.files["fundo"]
    planilha = request.files["planilha"]

    texto = request.form["texto"]
    tamanho = int(request.form["tamanho"])
    posicao = int(request.form["posicao"])

    id_lote = str(uuid.uuid4())[:8]

    pasta = os.path.join(PASTA_TEMP, id_lote)
    os.makedirs(pasta)

    caminho_fundo = os.path.join(pasta, fundo.filename)
    fundo.save(caminho_fundo)

    caminho_planilha = os.path.join(pasta, planilha.filename)
    planilha.save(caminho_planilha)

    df = pd.read_excel(caminho_planilha)

    coluna_nome = encontrar_coluna_nome(df)

    imagem_base = Image.open(caminho_fundo)

    largura, altura = imagem_base.size

    fonte = ImageFont.load_default()

    arquivos = []

    for i, row in df.iterrows():

        nome = str(row[coluna_nome]).strip()

        texto_final = texto.replace("{nome}", nome)

        img = imagem_base.copy()

        draw = ImageDraw.Draw(img)

        draw.text(
            (largura/2, altura/2 + posicao),
            texto_final,
            fill="black",
            font=fonte,
            anchor="mm",
            align="center"
        )

        nome_pdf = f"{nome}.pdf"

        caminho_pdf = os.path.join(pasta, nome_pdf)

        img.save(caminho_pdf, "PDF")

        arquivos.append(caminho_pdf)

    zip_nome = f"certificados_{id_lote}.zip"

    caminho_zip = os.path.join(pasta, zip_nome)

    with zipfile.ZipFile(
        caminho_zip,
        "w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=9
    ) as zipf:

        for arq in arquivos:

            zipf.write(
                arq,
                os.path.basename(arq)
            )

    return render_template(
        "download.html",
        arquivo=f"/download/{id_lote}/{zip_nome}"
    )


@app.route("/download/<lote>/<arquivo>")
def download(lote, arquivo):

    caminho = os.path.join(PASTA_TEMP, lote, arquivo)

    return send_file(
        caminho,
        as_attachment=True
    )


if __name__ == "__main__":
    app.run(debug=True)
