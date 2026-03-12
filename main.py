from flask import Flask, render_template, request, redirect, session, send_file
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile

app = Flask(__name__)
app.secret_key = "emitte_secret"

USER = "admin"
PASSWORD = "123"


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/guia")
def guia():
    return render_template("guia.html")


@app.route("/login", methods=["GET","POST"])
def login():

    if request.method == "POST":

        email = request.form["email"]
        senha = request.form["password"]

        if email == USER and senha == PASSWORD:
            session["user"] = email
            return redirect("/certificados")

    return render_template("login.html")


@app.route("/logout")
def logout():

    session.clear()
    return redirect("/")


@app.route("/certificados")
def certificados():

    if "user" not in session:
        return redirect("/login")

    return render_template("certificados.html")


def quebrar_linhas(texto, draw, font, largura_max):

    palavras = texto.split()
    linhas = []
    linha = ""

    for palavra in palavras:

        teste = linha + " " + palavra if linha else palavra
        largura = draw.textlength(teste, font=font)

        if largura <= largura_max:
            linha = teste
        else:
            linhas.append(linha)
            linha = palavra

    if linha:
        linhas.append(linha)

    return linhas


@app.route("/preview", methods=["POST"])
def preview():

    fundo = request.files["fundo"]
    texto = request.form["texto"]
    fonte_size = int(request.form["fonte"])
    alinhamento = request.form["alinhamento"]

    img = Image.open(fundo)

    largura, altura = img.size

    draw = ImageDraw.Draw(img)

    try:
        font = ImageFont.truetype("arial.ttf", fonte_size)
    except:
        font = ImageFont.load_default()

    largura_texto = largura * 0.7

    linhas = quebrar_linhas(texto, draw, font, largura_texto)

    y = altura * 0.45

    for linha in linhas:

        w = draw.textlength(linha, font=font)

        if alinhamento == "centro":
            x = (largura - w)/2
        elif alinhamento == "direita":
            x = largura - w - 100
        else:
            x = 100

        draw.text((x,y), linha, fill="black", font=font)

        y += fonte_size + 8

    buffer = io.BytesIO()

    img.save(buffer, format="PNG")

    buffer.seek(0)

    return send_file(buffer, mimetype="image/png")


@app.route("/gerar", methods=["POST"])
def gerar():

    planilha = request.files["planilha"]
    fundo = request.files["fundo"]

    texto = request.form["texto"]
    fonte_size = int(request.form["fonte"])
    alinhamento = request.form["alinhamento"]

    df = pd.read_excel(planilha)

    img_base = Image.open(fundo)

    largura, altura = img_base.size

    zip_buffer = io.BytesIO()
    zip_file = zipfile.ZipFile(zip_buffer,"w")

    try:
        font = ImageFont.truetype("arial.ttf", fonte_size)
    except:
        font = ImageFont.load_default()

    for _,row in df.iterrows():

        img = img_base.copy()
        draw = ImageDraw.Draw(img)

        texto_final = texto

        for col in df.columns:
            texto_final = texto_final.replace("{"+col+"}", str(row[col]))

        largura_texto = largura * 0.7

        linhas = quebrar_linhas(texto_final, draw, font, largura_texto)

        y = altura * 0.45

        for linha in linhas:

            w = draw.textlength(linha, font=font)

            if alinhamento == "centro":
                x = (largura - w)/2
            elif alinhamento == "direita":
                x = largura - w - 100
            else:
                x = 100

            draw.text((x,y), linha, fill="black", font=font)

            y += fonte_size + 8

        pdf_buffer = io.BytesIO()

        img.save(pdf_buffer, format="PDF")

        nome = str(row[df.columns[0]])

        zip_file.writestr(nome + ".pdf", pdf_buffer.getvalue())

    zip_file.close()

    zip_buffer.seek(0)

    return send_file(
        zip_buffer,
        download_name="certificados.zip",
        as_attachment=True
    )


if __name__ == "__main__":
    app.run()
