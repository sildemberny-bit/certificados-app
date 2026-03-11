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


def quebrar_texto(texto, largura=90):
    return textwrap.wrap(texto, largura)


def gerar_pdf(nome, curso, carga, texto, fundo_file,
              municipio, dia, mes, ano,
              fonte, alinhamento, posicao):

    buffer = io.BytesIO()

    c = canvas.Canvas(buffer, pagesize=landscape(A4))

    fundo = ImageReader(fundo_file)
    c.drawImage(fundo, 0, 0, width=842, height=595)

    texto_certificado = texto \
        .replace("{NOME}", str(nome)) \
        .replace("{CURSO}", str(curso)) \
        .replace("{CARGA}", str(carga))

    linhas = quebrar_texto(texto_certificado)

    if posicao == "acima":
        y = 360
    elif posicao == "abaixo":
        y = 260
    else:
        y = 310

    c.setFont("Helvetica", int(fonte))

    for linha in linhas:

        if alinhamento == "esquerda":
            c.drawString(120, y, linha)

        elif alinhamento == "direita":
            c.drawRightString(720, y, linha)

        else:
            c.drawCentredString(420, y, linha)

        y -= 28

    if municipio and dia and mes and ano:
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

        try:

            planilha = request.files.get("planilha")
            fundo = request.files.get("fundo")

            if not planilha or not fundo:
                return "Envie a planilha e o fundo do certificado."

            texto = request.form.get("texto", "")

            municipio = request.form.get("municipio", "")
            dia = request.form.get("dia", "")
            mes = request.form.get("mes", "")
            ano = request.form.get("ano", "")

            fonte = request.form.get("fonte", "16")
            alinhamento = request.form.get("alinhamento", "centro")
            posicao = request.form.get("posicao", "centro")

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
                        ano,
                        fonte,
                        alinhamento,
                        posicao
                    )

                    z.writestr(f"{nome}.pdf", pdf_buffer.read())

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

        fundo = request.files.get("fundo")

        if not fundo:
            return "Envie o fundo para gerar preview."

        texto = request.form.get("texto", "")

        municipio = request.form.get("municipio", "")
        dia = request.form.get("dia", "")
        mes = request.form.get("mes", "")
        ano = request.form.get("ano", "")

        fonte = request.form.get("fonte", "16")
        alinhamento = request.form.get("alinhamento", "centro")
        posicao = request.form.get("posicao", "centro")

        pdf_buffer = gerar_pdf(
            "NOME EXEMPLO",
            "CURSO EXEMPLO",
            "40",
            texto,
            fundo,
            municipio,
            dia,
            mes,
            ano,
            fonte,
            alinhamento,
            posicao
        )

        return send_file(
            pdf_buffer,
            download_name="preview.pdf"
        )

    except Exception as e:

        return f"Erro no preview: {str(e)}"


if __name__ == "__main__":
    app.run()
