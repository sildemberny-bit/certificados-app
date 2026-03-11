from flask import Flask, render_template, request, redirect, session, send_file
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.utils import ImageReader
import io
import zipfile

app = Flask(__name__)
app.secret_key = "emitte_secret"

USUARIO = "admin"
SENHA = "123"


def quebrar_linhas(texto, tamanho=80):
    linhas = []
    while texto:
        linhas.append(texto[:tamanho])
        texto = texto[tamanho:]
    return linhas


def gerar_pdf(nome, curso, carga, texto, fundo_file, municipio, dia, mes, ano):

    buffer = io.BytesIO()

    c = canvas.Canvas(buffer, pagesize=landscape(A4))

    fundo = ImageReader(fundo_file)
    c.drawImage(fundo, 0, 0, width=842, height=595)

    texto_certificado = texto \
        .replace("{NOME}", nome) \
        .replace("{CURSO}", curso) \
        .replace("{CARGA}", str(carga))

    linhas = quebrar_linhas(texto_certificado)

    y = 320

    for linha in linhas:
        c.setFont("Helvetica", 18)
        c.drawCentredString(420, y, linha)
        y -= 30

    if municipio:
        data_final = f"{municipio}, {dia} de {mes} de {ano}"
        c.setFont("Helvetica", 14)
        c.drawCentredString(420, 120, data_final)

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


@app.route("/certificados", methods=["GET", "POST"])
def certificados():

    if not session.get("logado"):
        return redirect("/")

    if request.method == "POST":

        planilha = request.files["planilha"]
        fundo = request.files["fundo"]

        texto = request.form.get("texto")

        municipio = request.form.get("municipio")
        dia = request.form.get("dia")
        mes = request.form.get("mes")
        ano = request.form.get("ano")

        df = pd.read_excel(planilha)

        memoria_zip = io.BytesIO()

        with zipfile.ZipFile(memoria_zip, "w") as z:

            for _, row in df.iterrows():

                nome = row["NOME"]
                curso = row["CURSO"]
                carga = row["CARGA"]

                fundo.seek(0)

                pdf_buffer = gerar_pdf(
                    nome,
                    curso,
                    carga,
                    texto,
                    fundo,
                    municipio,
                    dia,
                    mes,
                    ano
                )

                z.writestr(f"{nome}.pdf", pdf_buffer.read())

        memoria_zip.seek(0)

        return send_file(
            memoria_zip,
            download_name="certificados.zip",
            as_attachment=True
        )

    return render_template("certificados.html")


@app.route("/preview", methods=["POST"])
def preview():

    fundo = request.files["fundo"]
    texto = request.form.get("texto")

    municipio = request.form.get("municipio")
    dia = request.form.get("dia")
    mes = request.form.get("mes")
    ano = request.form.get("ano")

    pdf_buffer = gerar_pdf(
        "NOME EXEMPLO",
        "CURSO EXEMPLO",
        "40",
        texto,
        fundo,
        municipio,
        dia,
        mes,
        ano
    )

    return send_file(
        pdf_buffer,
        download_name="preview.pdf"
    )


if __name__ == "__main__":
    app.run()
