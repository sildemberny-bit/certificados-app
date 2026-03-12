from flask import Flask, render_template, request, send_file, redirect, session
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile

app = Flask(__name__)
app.secret_key = "emitte_secret"

USER = "admin"
PASSWORD = "123"


@app.route("/")
def home():
    return render_template("index.html")


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


@app.route("/preview", methods=["POST"])
def preview():

    try:

        fundo = request.files["fundo"]
        texto = request.form.get("texto","")
        fonte_size = int(request.form.get("fonte",12))
        alinhamento = request.form.get("alinhamento","centro")
        posicao = request.form.get("posicao","centro")

        img = Image.open(fundo).convert("RGB")

        largura, altura = img.size

        draw = ImageDraw.Draw(img)

        try:
            font = ImageFont.truetype("arial.ttf", fonte_size)
        except:
            font = ImageFont.load_default()

        largura_texto = int(largura * 0.8)

        linhas = []

        palavras = texto.split()

        linha = ""

        for palavra in palavras:

            teste = linha + " " + palavra

            w, h = draw.textsize(teste, font=font)

            if w <= largura_texto:

                linha = teste

            else:

                linhas.append(linha)

                linha = palavra

        linhas.append(linha)

        altura_bloco = len(linhas) * (fonte_size + 5)

        if posicao == "acima":
            y = int(altura * 0.35)
        elif posicao == "abaixo":
            y = int(altura * 0.65)
        else:
            y = int(altura * 0.5)

        for linha in linhas:

            w, h = draw.textsize(linha, font=font)

            if alinhamento == "centro":
                x = (largura - w) / 2
            elif alinhamento == "direita":
                x = largura - w - 100
            else:
                x = 100

            draw.text((x,y), linha, fill="black", font=font)

            y += fonte_size + 5


        buffer = io.BytesIO()

        img.save(buffer, format="PNG")

        buffer.seek(0)

        return send_file(buffer, mimetype="image/png")

    except Exception as e:

        return str(e), 500



@app.route("/gerar", methods=["POST"])
def gerar():

    try:

        planilha = request.files["planilha"]
        fundo = request.files["fundo"]

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


        for _, row in df.iterrows():

            img = img_base.copy()

            draw = ImageDraw.Draw(img)

            texto_final = texto

            for col in df.columns:
                texto_final = texto_final.replace("{"+col+"}", str(row[col]))

            largura_texto = int(largura * 0.8)

            palavras = texto_final.split()

            linhas = []

            linha = ""

            for palavra in palavras:

                teste = linha + " " + palavra

                w,h = draw.textsize(teste,font=font)

                if w <= largura_texto:
                    linha = teste
                else:
                    linhas.append(linha)
                    linha = palavra

            linhas.append(linha)

            if posicao == "acima":
                y = int(altura * 0.35)
            elif posicao == "abaixo":
                y = int(altura * 0.65)
            else:
                y = int(altura * 0.5)

            for linha in linhas:

                w,h = draw.textsize(linha,font=font)

                if alinhamento == "centro":
                    x = (largura - w)/2
                elif alinhamento == "direita":
                    x = largura - w - 100
                else:
                    x = 100

                draw.text((x,y), linha, fill="black", font=font)

                y += fonte_size + 5


            buffer = io.BytesIO()

            img.save(buffer, format="PDF")

            nome = str(row[df.columns[0]])

            zip_file.writestr(nome + ".pdf", buffer.getvalue())


        zip_file.close()

        zip_buffer.seek(0)

        return send_file(zip_buffer,
                         download_name="certificados.zip",
                         as_attachment=True)

    except Exception as e:

        return "Erro interno: " + str(e)


if __name__ == "__main__":
    app.run()
