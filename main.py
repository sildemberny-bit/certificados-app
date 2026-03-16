import os
import uuid
import zipfile
import pandas as pd

from flask import Flask, render_template, request, send_file
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


def otimizar_imagem(caminho):

    img = Image.open(caminho)

    novo_caminho = caminho.replace(".png", ".jpg").replace(".jpeg", ".jpg")

    img = img.convert("RGB")

    img.save(
        novo_caminho,
        "JPEG",
        quality=70,
        optimize=True
    )

    return novo_caminho


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/certificados")
def certificados_page():
    return render_template("certificados.html")


@app.route("/certificados", methods=["POST"])
def gerar_certificados():

    fundo = request.files["fundo"]
    planilha = request.files["planilha"]

    texto = request.form["texto"]
    fonte_tamanho = int(request.form["fonte"])

    alinhamento = request.form["alinhamento"]
    posicao_vertical = request.form["posicao_vertical"]

    ajuste_vertical = int(request.form["ajuste_vertical"])
    largura_texto = int(request.form["largura_texto"])

    id_lote = str(uuid.uuid4())[:8]

    pasta_lote = os.path.join(PASTA_TEMP, id_lote)
    os.makedirs(pasta_lote)

    caminho_fundo = os.path.join(pasta_lote, fundo.filename)
    fundo.save(caminho_fundo)

    caminho_fundo = otimizar_imagem(caminho_fundo)

    caminho_excel = os.path.join(pasta_lote, planilha.filename)
    planilha.save(caminho_excel)

    df = pd.read_excel(caminho_excel)

    coluna_nome = encontrar_coluna_nome(df)

    img_base = Image.open(caminho_fundo)

    largura_img, altura_img = img_base.size

    font = ImageFont.load_default()

    lista_pdfs = []

    for i, row in df.iterrows():

        nome = str(row[coluna_nome]).strip()

        texto_final = texto.replace("{nome}", nome)

        img = img_base.copy()
        draw = ImageDraw.Draw(img)

        largura_bloco = largura_img * (largura_texto / 100)

        if posicao_vertical == "superior":
            y = altura_img * 0.25
        elif posicao_vertical == "centro":
            y = altura_img * 0.5
        else:
            y = altura_img * 0.75

        y += ajuste_vertical

        x = largura_img / 2

        draw.multiline_text(
            (x, y),
            texto_final,
            fill="black",
            font=font,
            align="center",
            anchor="mm"
        )

        nome_pdf = f"{nome}.pdf"

        caminho_pdf = os.path.join(pasta_lote, nome_pdf)

        img.save(caminho_pdf, "PDF", resolution=100)

        lista_pdfs.append(caminho_pdf)

    zip_nome = f"certificados_{id_lote}.zip"

    caminho_zip = os.path.join(pasta_lote, zip_nome)

    with zipfile.ZipFile(
        caminho_zip,
        "w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=9
    ) as zipf:

        for arquivo in lista_pdfs:
            zipf.write(
                arquivo,
                os.path.basename(arquivo)
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
