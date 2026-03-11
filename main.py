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

        arquivo_excel = request.files["planilha"]
        fundo = request.files["fundo"]

        texto = request.form.get("texto")

        municipio = request.form.get("municipio")
        dia = request.form.get("dia")
        mes = request.form.get("mes")
        ano = request.form.get("ano")

        df = pd.read_excel(arquivo_excel)

        memoria_zip = io.BytesIO()

        with zipfile.ZipFile(memoria_zip, mode="w") as zf:

            for index, row in df.iterrows():

                buffer = io.BytesIO()

                c = canvas.Canvas(buffer, pagesize=landscape(A4))

                fundo_img = ImageReader(fundo)
                c.drawImage(fundo_img, 0, 0, width=842, height=595)

                nome = str(row["NOME"])
                curso = str(row["CURSO"])
                carga = str(row["CARGA"])

                texto_certificado = texto \
                    .replace("{NOME}", nome) \
                    .replace("{CURSO}", curso) \
                    .replace("{CARGA}", carga)

                largura = 700
                linhas = []

                while texto_certificado:
                    linhas.append(texto_certificado[:80])
                    texto_certificado = texto_certificado[80:]

                y = 320

                for linha in linhas:
                    c.setFont("Helvetica", 18)
                    c.drawCentredString(420, y, linha)
                    y -= 30

                data_final = f"{municipio}, {dia} de {mes} de {ano}"

                c.setFont("Helvetica", 14)
                c.drawCentredString(420, 120, data_final)

                c.save()

                buffer.seek(0)

                nome_pdf = f"{nome}.pdf"

                zf.writestr(nome_pdf, buffer.read())

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

    nome = "NOME EXEMPLO"
    curso = "CURSO EXEMPLO"
    carga = "40"

    texto_certificado = texto \
        .replace("{NOME}", nome) \
        .replace("{CURSO}", curso) \
        .replace("{CARGA}", carga)

    buffer = io.BytesIO()

    c = canvas.Canvas(buffer, pagesize=landscape(A4))

    fundo_img = ImageReader(fundo)
    c.drawImage(fundo_img, 0, 0, width=842, height=595)

    y = 320

    while texto_certificado:

        linha = texto_certificado[:80]
        texto_certificado = texto_certificado[80:]

        c.setFont("Helvetica", 18)
        c.drawCentredString(420, y, linha)
        y -= 30

    c.save()

    buffer.seek(0)

    return send_file(buffer, download_name="preview.pdf")


if __name__ == "__main__":
    app.run()
