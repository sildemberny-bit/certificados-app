from flask import Flask, render_template, request, redirect, session, send_file
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.utils import ImageReader
import io
import zipfile
import textwrap

app = Flask(__name__)
app.secret_key = "emitte_secret"

USUARIO = "admin"
SENHA = "123"


def substituir_variaveis(texto, linha):
    for coluna in linha.index:
        chave = "{" + coluna.upper() + "}"
        texto = texto.replace(chave, str(linha[coluna]))
    return texto


def quebrar_texto(texto, largura=85):
    return textwrap.wrap(texto, largura)


def gerar_pdf(linha, texto, fundo_file, fonte, alinhamento, posicao, municipio, dia, mes, ano):

    buffer = io.BytesIO()

    c = canvas.Canvas(buffer, pagesize=landscape(A4))

    fundo = ImageReader(fundo_file)
    c.drawImage(fundo, 0, 0, width=842, height=595)

    texto = substituir_variaveis(texto, linha)
    linhas = quebrar_texto(texto)

    if posicao == "acima":
        y = 360
    elif posicao == "abaixo":
        y = 250
    else:
        y = 310

    c.setFont("Helvetica", int(fonte))

    for linha_texto in linhas:

        if alinhamento == "esquerda":
            c.drawString(120, y, linha_texto)

        elif alinhamento == "direita":
            c.drawRightString(720, y, linha_texto)

        else:
            c.drawCentredString(420, y, linha_texto)

        y -= 28

    if municipio:
        data = f"{municipio}, {dia} de {mes} de {ano}"
        c.setFont("Helvetica", 14)
        c.drawCentredString(420, 120, data)

    c.save()

    buffer.seek(0)
    return buffer


@app.route("/", methods=["GET", "POST"])
def login():

    if request.method == "POST":

        usuario = request.form.get("usuario")
        senha = request.form.get("senha")

        if usuario == USUARIO and senha == SENHA:
            session["logado"] = True
            return redirect("/certificados")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")


@app.route("/certificados", methods=["GET", "POST"])
def certificados():

    if not session.get("logado"):
        return redirect("/")

    if request.method == "POST":

        try:

            planilha = request.files["planilha"]
            fundo = request.files["fundo"]

            texto = request.form.get("texto")

            fonte = request.form.get("fonte")
            alinhamento = request.form.get("alinhamento")
            posicao = request.form.get("posicao")

            municipio = request.form.get("municipio")
            dia = request.form.get("dia")
            mes = request.form.get("mes")
            ano = request.form.get("ano")

            df = pd.read_excel(planilha)

            memoria_zip = io.BytesIO()

            with zipfile.ZipFile(memoria_zip, "w") as z:

                for _, linha in df.iterrows():

                    fundo.seek(0)

                    pdf = gerar_pdf(
                        linha,
                        texto,
                        fundo,
                        fonte,
                        alinhamento,
                        posicao,
                        municipio,
                        dia,
                        mes,
                        ano
                    )

                    nome = str(linha.iloc[0]).replace(" ", "_")

                    z.writestr(f"{nome}.pdf", pdf.read())

            memoria_zip.seek(0)

            return send_file(
                memoria_zip,
                download_name="certificados.zip",
                as_attachment=True
            )

        except Exception as e:

            return f"Erro interno: {str(e)}"

    return render_template("certificados.html")


@app.route("/preview", methods=["POST"])
def preview():

    try:

        fundo = request.files["fundo"]

        texto = request.form.get("texto")

        fonte = request.form.get("fonte")
        alinhamento = request.form.get("alinhamento")
        posicao = request.form.get("posicao")

        municipio = request.form.get("municipio")
        dia = request.form.get("dia")
        mes = request.form.get("mes")
        ano = request.form.get("ano")

        exemplo = pd.Series({
            "NOME": "Nome Exemplo",
            "PROJETO": "Projeto de Demonstração"
        })

        pdf = gerar_pdf(
            exemplo,
            texto,
            fundo,
            fonte,
            alinhamento,
            posicao,
            municipio,
            dia,
            mes,
            ano
        )

        return send_file(pdf, download_name="preview.pdf")

    except Exception as e:
        return f"Erro preview: {str(e)}"


if __name__ == "__main__":
    app.run()
