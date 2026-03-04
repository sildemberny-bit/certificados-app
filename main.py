from flask import Flask, render_template, request, send_file, redirect
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import os
import zipfile

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "certificados"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route("/")
def home():
    return redirect("/certificados")

@app.route("/certificados", methods=["GET", "POST"])
def certificados():
    if request.method == "POST":
        fundo = request.files["fundo"]
        planilha = request.files["planilha"]
        texto_modelo = request.form["texto"]

        fundo_path = os.path.join(UPLOAD_FOLDER, fundo.filename)
        planilha_path = os.path.join(UPLOAD_FOLDER, planilha.filename)

        fundo.save(fundo_path)
        planilha.save(planilha_path)

        df = pd.read_excel(planilha_path)

        # 🔥 PADRONIZA TODAS COLUNAS PARA MAIÚSCULAS
        df.columns = df.columns.str.upper()

        arquivos_gerados = []

        for index, row in df.iterrows():

            texto_final = texto_modelo

            # 🔥 SUBSTITUI AUTOMATICAMENTE TODOS OS CAMPOS
            for coluna in df.columns:
                valor = str(row[coluna])
                texto_final = texto_final.replace("{" + coluna + "}", valor)

            nome_arquivo = f"certificado_{index}.pdf"
            caminho_pdf = os.path.join(OUTPUT_FOLDER, nome_arquivo)

            c = canvas.Canvas(caminho_pdf, pagesize=A4)
            largura, altura = A4

            # Fundo
            c.drawImage(fundo_path, 0, 0, width=largura, height=altura)

            # Texto centralizado
            c.setFont("Helvetica", 18)
            c.drawCentredString(largura / 2, altura / 2, texto_final)

            c.save()

            arquivos_gerados.append(caminho_pdf)

        # Criar ZIP com todos certificados
        zip_path = os.path.join(OUTPUT_FOLDER, "certificados.zip")

        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for arquivo in arquivos_gerados:
                zipf.write(arquivo, os.path.basename(arquivo))

        return send_file(zip_path, as_attachment=True)

    return render_template("certificados.html")

if __name__ == "__main__":
    app.run(debug=True)
