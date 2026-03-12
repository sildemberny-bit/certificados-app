from flask import Flask, render_template, request, send_file, redirect, session
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile

app = Flask(__name__)
app.secret_key = "segredo"

EMAIL = "admin"
SENHA = "123"


@app.route("/", methods=["GET","POST"])
def login():

    if request.method == "POST":

        email = request.form["email"]
        password = request.form["password"]

        if email == EMAIL and password == SENHA:
            session["user"] = email
            return redirect("/certificados")

    return render_template("login.html")


@app.route("/certificados")
def certificados():

    if "user" not in session:
        return redirect("/")

    return render_template("certificados.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")


@app.route("/preview", methods=["POST"])
def preview():

    fundo = request.files["fundo"]
    texto = request.form["texto"]
    fonte_size = int(request.form["fonte"])
    alinhamento = request.form["alinhamento"]

    img = Image.open(fundo.stream).convert("RGB")
    largura, altura = img.size

    draw = ImageDraw.Draw(img)

    try:
        fonte = ImageFont.truetype("arial.ttf", fonte_size)
    except:
        fonte = ImageFont.load_default()

    y = altura * 0.45

    linhas = texto.split("\n")

    for linha in linhas:

        w,h = draw.textsize(linha,font=fonte)

        if alinhamento == "centro":
            x = (largura-w)/2
        elif alinhamento == "direita":
            x = largura-w-100
        else:
            x = 100

        draw.text((x,y),linha,fill="black",font=fonte)

        y += h + 10

    buffer = io.BytesIO()
    img.save(buffer,"PNG")
    buffer.seek(0)

    return send_file(buffer,mimetype="image/png")


@app.route("/gerar", methods=["POST"])
def gerar():

    planilha = request.files["planilha"]
    fundo = request.files["fundo"]

    texto = request.form["texto"]
    fonte_size = int(request.form["fonte"])
    alinhamento = request.form["alinhamento"]

    df = pd.read_excel(planilha)

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer,"w") as zipf:

        for i,row in df.iterrows():

            img = Image.open(fundo.stream).convert("RGB")
            largura, altura = img.size

            draw = ImageDraw.Draw(img)

            try:
                fonte = ImageFont.truetype("arial.ttf", fonte_size)
            except:
                fonte = ImageFont.load_default()

            texto_cert = texto

            for coluna in df.columns:
                texto_cert = texto_cert.replace("{"+coluna+"}", str(row[coluna]))

            y = altura * 0.45

            linhas = texto_cert.split("\n")

            for linha in linhas:

                w,h = draw.textsize(linha,font=fonte)

                if alinhamento == "centro":
                    x = (largura-w)/2
                elif alinhamento == "direita":
                    x = largura-w-100
                else:
                    x = 100

                draw.text((x,y),linha,fill="black",font=fonte)

                y += h + 10

            buffer = io.BytesIO()
            img.save(buffer,"PDF")
            buffer.seek(0)

            nome = str(row[df.columns[0]])+".pdf"

            zipf.writestr(nome,buffer.read())

    zip_buffer.seek(0)

    return send_file(
        zip_buffer,
        download_name="certificados.zip",
        as_attachment=True
    )


@app.route("/guia")
def guia():
    return render_template("guia.html")


if __name__ == "__main__":
    app.run()
