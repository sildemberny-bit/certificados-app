from flask import Flask, render_template, request, send_file, redirect, session
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
import os
import zipfile
import textwrap
import re

app = Flask(__name__)
app.secret_key = "emitte_super_secreta"

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "certificados"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

USUARIO_LOGIN = "admin"
USUARIO_SENHA = "123"

# ===== LOGIN =====
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        login = request.form["email"]
        senha = request.form["password"]

        if login == USUARIO_LOGIN and senha == USUARIO_SENHA:
            session["usuario"] = login
            return redirect("/certificados")

    return render_template("login.html")

# ===== LOGOUT =====
@app.route("/logout")
def logout():
    session.pop("usuario", None)
    return redirect("/login")

# ===== HOME =====
@app.route("/")
def home():
    return redirect("/login")

# ===== CERTIFICADOS =====
@app.route("/certificados", methods=["GET", "POST"])
def certificados():

    if "usuario" not in session:
        return redirect("/login")

    if request.method == "POST":
        fundo = request.files["fundo"]
        planilha = request.files["planilha"]
        texto_modelo = request.form["texto"]
        tamanho_fonte = int(request.form["fonte"])
        alinhamento = request.form["alinhamento"]
        posicao_vertical = request.form["posicao_vertical"]

        fundo_path = os.path.join(UPLOAD_FOLDER, fundo.filename)
        planilha_path = os.path.join(UPLOAD_FOLDER, planilha.filename)

        fundo.save(fundo_path)
        planilha.save(planilha_path)

        df = pd.read_excel(planilha_path)

        # Normaliza colunas
        df.columns = df.columns.str.strip()

        arquivos_gerados = []

        for index, row in df.iterrows():

            texto_final = texto_modelo

            # Procura todos os campos {qualquer_coisa}
            campos = re.findall(r"\{(.*?)\}", texto_modelo)

            for campo in campos:
                for coluna in df.columns:
                    if campo.strip().lower() == coluna.strip().lower():
                        valor = str(row[coluna])
                        texto_final = re.sub(
                            r"\{" + campo + r"\}",
                            valor,
                            texto_final,
                            flags=re.IGNORECASE
                        )

            nome_arquivo = f"certificado_{index}.pdf"
            caminho_pdf = os.path.join(OUTPUT_FOLDER, nome_arquivo)

            largura, altura = landscape(A4)
            c = canvas.Canvas(caminho_pdf, pagesize=(largura, altura))

            fundo_img = ImageReader(fundo_path)
            c.drawImage(fundo_img, 0, 0, width=largura, height=altura)

            c.setFont("Helvetica", tamanho_fonte)

            linhas = textwrap.wrap(texto_final, width=80)

            if posicao_vertical == "superior":
                y_inicial = altura * 0.65
            elif posicao_vertical == "inferior":
                y_inicial = altura * 0.35
            else:
                y_inicial = altura * 0.50

            espacamento = tamanho_fonte + 8

            for i, linha in enumerate(linhas):
                y = y_inicial - (i * espacamento)

                if alinhamento == "centro":
                    c.drawCentredString(largura / 2, y, linha)
                elif alinhamento == "esquerda":
                    c.drawString(100, y, linha)
                elif alinhamento == "direita":
                    c.drawRightString(largura - 100, y, linha)

            c.save()
            arquivos_gerados.append(caminho_pdf)

        zip_path = os.path.join(OUTPUT_FOLDER, "certificados.zip")

        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for arquivo in arquivos_gerados:
                zipf.write(arquivo, os.path.basename(arquivo))

        return send_file(zip_path, as_attachment=True)

    return render_template("certificados.html")

if __name__ == "__main__":
    app.run(debug=True)
