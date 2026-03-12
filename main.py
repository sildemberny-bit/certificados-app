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


# =========================
# LANDING
# =========================

@app.route("/")
def landing():
    return render_template("landing.html")


@app.route("/guia")
def guia():
    return render_template("guia.html")


# =========================
# LOGIN
# =========================

@app.route("/login", methods=["GET", "POST"])
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


# =========================
# FUNÇÕES AUXILIARES
# =========================

def substituir_variaveis(texto, linha):

    for coluna in linha.index:
        chave = "{" + coluna.upper() + "}"
        texto = texto.replace(chave, str(linha[coluna]))

    return texto


# quebra texto usando largura maior
def quebrar_texto(texto):
    return textwrap.wrap(texto, 140)


# =========================
# GERAR PDF
# =========================

def gerar_pdf(linha, texto, fundo_bytes, fonte, alinhamento, posicao, municipio, dia, mes, ano):

    buffer = io.BytesIO()

    c = canvas.Canvas(buffer, pagesize=landscape(A4))

    fundo = ImageReader(io.BytesIO(fundo_bytes))

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

        y -= 30

    if municipio:
        data = f"{municipio}, {dia} de {mes} de {ano}"
        c.setFont("Helvetica", 14)
        c.drawCentredString(420, 120, data)

    c.save()

    buffer.seek(0)

    return buffer


# =========================
# GERAR CERTIFICADOS
# =========================

@app.route("/certificados", methods=["GET", "POST"])
def certificados():

    if not session.get("logado"):
        return redirect("/login")

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

            fundo_bytes = fundo.read()

            memoria_zip = io.BytesIO()

            with zipfile.ZipFile(memoria_zip, "w") as z:

                for _, linha in df.iterrows():

                    pdf = gerar_pdf(
                        linha,
                        texto,
                        fundo_bytes,
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


# =========================
# PREVIEW
# =========================

@app.route("/preview", methods=["POST"])
def preview():

    try:

        fundo = request.files.get("fundo")

        if not fundo:
            return ""

        fundo_bytes = fundo.read()

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
            "PROJETO": "Projeto Demonstração"
        })

        pdf = gerar_pdf(
            exemplo,
            texto,
            fundo_bytes,
            fonte,
            alinhamento,
            posicao,
            municipio,
            dia,
            mes,
            ano
        )

        return send_file(pdf, mimetype="application/pdf")

    except Exception as e:

        return f"Erro preview: {str(e)}"


# =========================
# RUN
# =========================

if __name__ == "__main__":
    app.run()
