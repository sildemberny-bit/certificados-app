from flask import Flask, render_template, request, send_file, redirect, session
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
import os
import zipfile
import re
from datetime import datetime
from reportlab.pdfbase.pdfmetrics import stringWidth

app = Flask(__name__)
app.secret_key = "emitte_super_secreta"

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "certificados"
METRICS_FILE = "metrics.txt"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

USUARIO_LOGIN = "admin"
USUARIO_SENHA = "123"


# ===== MÉTRICAS =====

def registrar_metricas(qtd_certificados):

    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if not os.path.exists(METRICS_FILE):
        with open(METRICS_FILE, "w") as f:
            f.write("TOTAL_LOTES=0\n")
            f.write("TOTAL_CERTIFICADOS=0\n")

    with open(METRICS_FILE, "r") as f:
        linhas = f.readlines()

    total_lotes = int(linhas[0].split("=")[1])
    total_certificados = int(linhas[1].split("=")[1])

    total_lotes += 1
    total_certificados += qtd_certificados

    with open(METRICS_FILE, "w") as f:
        f.write(f"TOTAL_LOTES={total_lotes}\n")
        f.write(f"TOTAL_CERTIFICADOS={total_certificados}\n")

    print("===== NOVA GERAÇÃO REALIZADA =====")
    print(f"Data/Hora: {agora}")
    print(f"Lote atual: {total_lotes}")
    print(f"Certificados neste lote: {qtd_certificados}")
    print(f"Total acumulado certificados: {total_certificados}")
    print("===================================")


# ===== LANDING =====

@app.route("/")
def landing():
    return render_template("landing.html")


# ===== GUIA =====

@app.route("/guia")
def guia():
    return render_template("guia.html")


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
    return redirect("/")


# ===== FUNÇÃO DE QUEBRA DE TEXTO INTELIGENTE =====

def quebrar_linhas(texto, largura_max, fonte_nome, fonte_tamanho):

    palavras = texto.split()
    linhas = []
    linha_atual = ""

    for palavra in palavras:

        teste = linha_atual + " " + palavra if linha_atual else palavra

        largura = stringWidth(teste, fonte_nome, fonte_tamanho)

        if largura <= largura_max:
            linha_atual = teste
        else:
            linhas.append(linha_atual)
            linha_atual = palavra

    if linha_atual:
        linhas.append(linha_atual)

    return linhas


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
        df.columns = df.columns.str.strip()

        arquivos_gerados = []

        for index, row in df.iterrows():

            texto_final = texto_modelo

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

            texto_final = re.sub(r"\.\s+([a-z])", lambda m: ". " + m.group(1).upper(), texto_final)

            nome_arquivo = f"certificado_{index}.pdf"
            caminho_pdf = os.path.join(OUTPUT_FOLDER, nome_arquivo)

            largura, altura = landscape(A4)

            c = canvas.Canvas(caminho_pdf, pagesize=(largura, altura))

            fundo_img = ImageReader(fundo_path)
            c.drawImage(fundo_img, 0, 0, width=largura, height=altura)

            fonte_nome = "Helvetica"

            c.setFont(fonte_nome, tamanho_fonte)

            margem = largura * 0.125
            largura_texto = largura - (margem * 2)

            paragrafos = texto_final.split("\n\n")

            linhas = []

            for p in paragrafos:

                linhas_paragrafo = quebrar_linhas(p, largura_texto, fonte_nome, tamanho_fonte)

                linhas.extend(linhas_paragrafo)
                linhas.append("")

            espacamento = tamanho_fonte * 1.35

            if posicao_vertical == "superior":
                y_inicial = altura * 0.65
            elif posicao_vertical == "inferior":
                y_inicial = altura * 0.35
            else:
                y_inicial = altura * 0.55

            y = y_inicial

            for linha in linhas:

                if linha == "":
                    y -= espacamento
                    continue

                if alinhamento == "centro":

                    c.drawCentredString(largura / 2, y, linha)

                elif alinhamento == "esquerda":

                    c.drawString(margem, y, linha)

                elif alinhamento == "direita":

                    c.drawRightString(largura - margem, y, linha)

                elif alinhamento == "justificado":

                    palavras = linha.split()

                    if len(palavras) == 1:
                        c.drawString(margem, y, linha)

                    else:

                        largura_palavras = sum(
                            stringWidth(p, fonte_nome, tamanho_fonte)
                            for p in palavras
                        )

                        espacos = len(palavras) - 1

                        espaco_extra = (largura_texto - largura_palavras) / espacos

                        x = margem

                        for palavra in palavras:

                            c.drawString(x, y, palavra)

                            x += stringWidth(palavra, fonte_nome, tamanho_fonte) + espaco_extra

                y -= espacamento

            c.save()

            arquivos_gerados.append(caminho_pdf)

        zip_path = os.path.join(OUTPUT_FOLDER, "certificados.zip")

        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for arquivo in arquivos_gerados:
                zipf.write(arquivo, os.path.basename(arquivo))

        registrar_metricas(len(arquivos_gerados))

        return send_file(zip_path, as_attachment=True)

    return render_template("certificados.html")


if __name__ == "__main__":
    app.run(debug=True)
