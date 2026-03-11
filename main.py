from flask import Flask, render_template, request, send_file, redirect
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
import os
import zipfile

app = Flask(__name__)

USUARIO = "admin"
SENHA = "123"

contador_certificados = 0


@app.route("/", methods=["GET","POST"])
def login():

    if request.method == "POST":

        usuario = request.form["usuario"]
        senha = request.form["senha"]

        if usuario == USUARIO and senha == SENHA:
            return redirect("/certificados")

    return render_template("login.html")


@app.route("/certificados", methods=["GET","POST"])
def certificados():

    global contador_certificados

    if request.method == "POST":

        arquivo_excel = request.files["excel"]
        fundo = request.files["fundo"]
        texto = request.form["texto"]

        tamanho = int(request.form["tamanho"])
        alinhamento = request.form["alinhamento"]

        municipio = request.form["municipio"]
        dia = request.form["dia"]
        mes = request.form["mes"]
        ano = request.form["ano"]

        data = f"{dia} de {mes} de {ano}"

        df = pd.read_excel(arquivo_excel)

        os.makedirs("temp", exist_ok=True)

        arquivos_pdf = []

        largura, altura = landscape(A4)

        fundo_path = "temp/fundo.jpg"
        fundo.save(fundo_path)

        for i, linha in df.iterrows():

            conteudo = texto

            for coluna in df.columns:
                conteudo = conteudo.replace(
                    "{"+coluna+"}",
                    str(linha[coluna])
                )

            conteudo = conteudo.replace("{MUNICIPIO}", municipio)
            conteudo = conteudo.replace("{DIA}", dia)
            conteudo = conteudo.replace("{MES}", mes)
            conteudo = conteudo.replace("{ANO}", ano)
            conteudo = conteudo.replace("{DATA}", data)

            nome_pdf = str(linha[df.columns[0]]) + ".pdf"

            caminho_pdf = "temp/" + nome_pdf

            c = canvas.Canvas(caminho_pdf, pagesize=landscape(A4))

            c.drawImage(fundo_path,0,0,width=largura,height=altura)

            y = altura/2

            largura_texto = largura - 6*cm

            linhas = []

            palavras = conteudo.split()

            linha_atual = ""

            for palavra in palavras:

                teste = linha_atual + " " + palavra

                if c.stringWidth(teste,"Helvetica",tamanho) < largura_texto:
                    linha_atual = teste
                else:
                    linhas.append(linha_atual)
                    linha_atual = palavra

            linhas.append(linha_atual)

            for linha_texto in linhas:

                if alinhamento == "centro":
                    c.drawCentredString(largura/2,y,linha_texto)

                elif alinhamento == "esquerda":
                    c.drawString(3*cm,y,linha_texto)

                elif alinhamento == "direita":
                    c.drawRightString(largura-3*cm,y,linha_texto)

                y -= tamanho + 5

            c.save()

            arquivos_pdf.append(caminho_pdf)

        zip_path = "certificados.zip"

        with zipfile.ZipFile(zip_path,"w") as zipf:

            for pdf in arquivos_pdf:
                zipf.write(pdf, os.path.basename(pdf))

        contador_certificados += len(arquivos_pdf)

        return send_file(zip_path, as_attachment=True)

    return render_template("certificados.html")


@app.route("/contador")
def contador():
    global contador_certificados
    return f"Certificados gerados: {contador_certificados}"


@app.route("/guia")
def guia():
    return render_template("guia.html")


if __name__ == "__main__":
    app.run()
