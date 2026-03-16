from flask import Flask, render_template, request, send_file
import pandas as pd
import os
import uuid
import zipfile
from PIL import Image, ImageDraw, ImageFont

app = Flask(__name__)

TEMP_DIR = "temp"

if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/certificados")
def certificados():
    return render_template("certificados.html")


@app.route("/gerar", methods=["POST"])
def gerar():

    fundo = request.files["fundo"]
    planilha = request.files["planilha"]

    texto = request.form["texto"]
    fonte_tamanho = int(request.form["fonte"])
    posicao_vertical = int(request.form["posicao"])

    lote = str(uuid.uuid4())[:8]

    pasta = os.path.join(TEMP_DIR, lote)
    os.makedirs(pasta)

    caminho_fundo = os.path.join(pasta, fundo.filename)
    fundo.save(caminho_fundo)

    caminho_planilha = os.path.join(pasta, planilha.filename)
    planilha.save(caminho_planilha)

    df = pd.read_excel(caminho_planilha)

    if "nome" not in [c.lower() for c in df.columns]:
        raise Exception("Planilha precisa ter coluna 'nome'")

    coluna_nome = [c for c in df.columns if c.lower() == "nome"][0]

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
            (largura/2, altura/2 + posicao_vertical),
            texto_final,
            fill="black",
            font=fonte,
            anchor="mm",
            align="center"
        )

        pdf_nome = f"{nome}.pdf"

        caminho_pdf = os.path.join(pasta, pdf_nome)

        img.save(caminho_pdf, "PDF")

        arquivos.append(caminho_pdf)

    zip_nome = f"certificados_{lote}.zip"

    zip_path = os.path.join(pasta, zip_nome)

    # 🔹 compressão máxima do zip
    with zipfile.ZipFile(
        zip_path,
        "w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=9
    ) as zipf:

        for arq in arquivos:
            zipf.write(arq, os.path.basename(arq))

    return render_template(
        "download.html",
        arquivo=f"/download/{lote}/{zip_nome}"
    )


@app.route("/download/<lote>/<arquivo>")
def download(lote, arquivo):

    caminho = os.path.join(TEMP_DIR, lote, arquivo)

    return send_file(
        caminho,
        as_attachment=True
    )


if __name__ == "__main__":
    app.run(debug=True)
