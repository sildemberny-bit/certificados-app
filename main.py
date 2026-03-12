from flask import Flask, render_template, request, redirect, session, send_file
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile

app = Flask(__name__)
app.secret_key = "emitte_secret"

USER = "admin"
PASSWORD = "123"


# LANDING
@app.route("/")
def index():
    return render_template("index.html")


# LOGIN
@app.route("/login", methods=["GET","POST"])
def login():

    if request.method == "POST":

        email = request.form.get("email")
        senha = request.form.get("password")

        if email == USER and senha == PASSWORD:
            session["user"] = email
            return redirect("/certificados")

    return render_template("login.html")


# LOGOUT
@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")


# TELA PRINCIPAL
@app.route("/certificados")
def certificados():

    if "user" not in session:
        return redirect("/login")

    return render_template("certificados.html")


# QUEBRA DE LINHAS
def quebrar_linhas(texto, draw, font, largura_max):

    palavras = texto.split()
    linhas = []
    linha = ""

    for palavra in palavras:

        teste = linha + " " + palavra if linha else palavra
        largura, _ = draw.textbbox((0,0), teste, font=font)[2:]

        if largura <= largura_max:
            linha = teste
        else:
            linhas.append(linha)
            linha = palavra

    if linha:
        linhas.append(linha)

    return linhas


# PREVIEW
@app.route("/preview", methods=["POST"])
def preview():

    try:

        fundo = request.files.get("fundo")
        texto = request.form.get("texto","")
        fonte_size = int(request.form.get("fonte",12))
        alinhamento = request.form.get("alinhamento","centro")
        posicao = request.form.get("posicao","centro")

        if not fundo:
            return "sem fundo",400

        img = Image.open(fundo).convert("RGB")

        largura, altura = img.size

        draw = ImageDraw.Draw(img)

        try:
            font = ImageFont.truetype("arial.ttf", fonte_size)
        except:
            font = ImageFont.load_default()

        largura_texto = int(largura * 0.75)

        linhas = quebrar_linhas(texto, draw, font, largura_texto)

        if posicao == "acima":
            y = int(altura * 0.35)
        elif posicao == "abaixo":
            y = int(altura * 0.65)
        else:
            y = int(altura * 0.5)

        for linha in linhas:

            w = draw.textbbox((0,0), linha, font=font)[2]

            if alinhamento == "centro":
                x = (largura - w)//2
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

    except Exception as e:
        return str(e),500


# GERAR CERTIFICADOS
@app.route("/gerar", methods=["POST"])
def gerar():

    try:

        planilha = request.files.get("planilha")
        fundo = request.files.get("fundo")

        texto = request.form.get("texto","")
        fonte_size = int(request.form.get("fonte",12))
        alinhamento = request.form.get("alinhamento","centro")
        posicao = request.form.get("posicao","centro")

        df = pd.read_excel(planilha)

        img_base = Image.open(fundo).convert("RGB")

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

            largura_texto = int(largura * 0.75)

            linhas = quebrar_linhas(texto_final, draw, font, largura_texto)

            if posicao == "acima":
                y = int(altura * 0.35)
            elif posicao == "abaixo":
                y = int(altura * 0.65)
            else:
                y = int(altura * 0.5)

            for linha in linhas:

                w = draw.textbbox((0,0), linha, font=font)[2]

                if alinhamento == "centro":
                    x = (largura - w)//2
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

    except Exception as e:
        return "Erro interno: " + str(e)


if __name__ == "__main__":
    app.run()
