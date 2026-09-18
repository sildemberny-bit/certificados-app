from flask import Flask, render_template, request, redirect, send_file, session
import pandas as pd
from PIL import Image, ImageOps
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import zipfile
import os
import unicodedata
import re
import tempfile
import datetime
import shutil


app = Flask(__name__)
app.secret_key = "emitte_secret"


USUARIO = "admin"
SENHA = "123"


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PASTA_DOWNLOAD = os.path.join(BASE_DIR, "downloads")

os.makedirs(PASTA_DOWNLOAD, exist_ok=True)


# ============================================================
# ROTAS BÁSICAS
# ============================================================

@app.route("/")
def landing():
    return render_template("landing.html")


@app.route("/guia")
def guia():
    return render_template("guia.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = request.form.get("email")
        senha = request.form.get("password")

        if usuario == USUARIO and senha == SENHA:
            session["user"] = usuario
            return redirect("/certificados")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.pop("user", None)
    return redirect("/login")


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def limpar_nome_arquivo(nome):
    """
    Converte o nome do participante em um nome seguro para arquivo.
    """
    nome = str(nome).lower()

    nome = unicodedata.normalize("NFD", nome)
    nome = nome.encode("ascii", "ignore").decode("utf-8")

    nome = re.sub(r"[^a-z0-9 ]", "", nome)
    nome = nome.replace(" ", "_")

    return nome


def substituir_campos(texto, linha):
    """
    Substitui campos como:
        {nome}
        {NOME}
        {Nome}

    pelos valores existentes na planilha.
    """

    resultado = texto

    for coluna in linha.index:
        valor = str(linha[coluna])

        resultado = resultado.replace(
            "{" + str(coluna) + "}",
            f"<b>{valor}</b>"
        )

        resultado = resultado.replace(
            "{" + str(coluna).lower() + "}",
            f"<b>{valor}</b>"
        )

        resultado = resultado.replace(
            "{" + str(coluna).upper() + "}",
            f"<b>{valor}</b>"
        )

    return resultado


def criar_estilo(fonte, alinhamento):
    """
    Cria o estilo utilizado no texto do certificado.
    """

    if alinhamento == "centro":
        alinh = TA_CENTER

    elif alinhamento == "esquerda":
        alinh = TA_LEFT

    else:
        alinh = TA_RIGHT

    return ParagraphStyle(
        name="Certificado",
        fontName="Helvetica",
        fontSize=fonte,
        leading=fonte * 1.4,
        alignment=alinh
    )


def preparar_imagem_fundo(fundo):
    """
    Abre a imagem enviada pelo usuário, corrige orientação EXIF
    e reduz imagens excessivamente grandes.

    Retorna uma imagem PIL pronta para ser usada pelo ImageReader.

    A imagem é preparada UMA ÚNICA VEZ.
    """

    largura_pagina, altura_pagina = landscape(A4)

    # 150 DPI.
    # A4 paisagem fica aproximadamente em:
    # 1754 x 1241 pixels.
    dpi = 150

    largura_alvo = int((largura_pagina / 72) * dpi)
    altura_alvo = int((altura_pagina / 72) * dpi)

    imagem = Image.open(fundo)

    # Corrige orientação proveniente de celulares/câmeras.
    imagem = ImageOps.exif_transpose(imagem)

    # Proteção contra arquivos absurdamente grandes.
    largura_original, altura_original = imagem.size
    pixels = largura_original * altura_original

    limite_pixels = 40_000_000

    if pixels > limite_pixels:
        imagem.close()

        raise ValueError(
            "A imagem enviada possui dimensões muito grandes. "
            "Reduza a resolução da imagem e tente novamente."
        )

    # Converte para RGB.
    if imagem.mode != "RGB":
        imagem = imagem.convert("RGB")

    # Redimensiona uma única vez.
    #
    # O sistema antigo esticava a imagem para ocupar toda a página.
    # Mantemos esse comportamento para preservar a aparência esperada.
    imagem = imagem.resize(
        (largura_alvo, altura_alvo),
        Image.Resampling.LANCZOS
    )

    return imagem


def gerar_pdf_lote(
    imagem,
    df,
    texto,
    fonte,
    alinhamento,
    posicao_vertical,
    ajuste_vertical,
    largura_texto_percent,
    caminho_pdf
):
    """
    Gera UM ÚNICO PDF contendo todos os certificados.

    Esta é a arquitetura que já funcionou anteriormente:
    
        uma Canvas
        uma imagem de fundo
        várias páginas
    """

    largura_pagina, altura_pagina = landscape(A4)

    c = canvas.Canvas(
        caminho_pdf,
        pagesize=(largura_pagina, altura_pagina)
    )

    largura_texto = (
        largura_pagina *
        (largura_texto_percent / 100)
    )

    style = criar_estilo(
        fonte,
        alinhamento
    )

    # IMPORTANTE:
    # O ImageReader é criado UMA ÚNICA VEZ.
    #
    # O mesmo objeto é reutilizado nas páginas do documento.
    fundo_reader = ImageReader(imagem)

    for _, linha in df.iterrows():

        texto_certificado = substituir_campos(
            texto,
            linha
        )

        texto_certificado = texto_certificado.replace(
            "\n",
            "<br/>"
        )

        # Fundo do certificado.
        c.drawImage(
            fundo_reader,
            0,
            0,
            width=largura_pagina,
            height=altura_pagina
        )

        # Texto.
        p = Paragraph(
            texto_certificado,
            style
        )

        w, h = p.wrap(
            largura_texto,
            altura_pagina
        )

        # Posicionamento vertical.
        if posicao_vertical == "superior":

            y = altura_pagina * 0.75

        elif posicao_vertical == "centro":

            y = (
                (altura_pagina / 2)
                - (h / 2)
            )

        else:

            y = altura_pagina * 0.30

        y = y + ajuste_vertical

        # Centralização horizontal da caixa de texto.
        x = (
            largura_pagina - largura_texto
        ) / 2

        p.drawOn(
            c,
            x,
            y
        )

        c.showPage()

    c.save()


def detectar_coluna_nome(df):
    """
    Procura automaticamente a coluna que contém o nome.
    """

    for col in df.columns:

        col_norm = unicodedata.normalize(
            "NFD",
            str(col)
        )

        col_norm = (
            col_norm
            .encode("ascii", "ignore")
            .decode("utf-8")
        )

        col_norm = col_norm.lower()

        if "nome" in col_norm:
            return col

    # Se não encontrar uma coluna com "nome",
    # utiliza a primeira coluna da planilha.
    return df.columns[0]


def dividir_pdf(caminho_pdf, df, pasta_saida):
    """
    Divide o PDF multipágina em PDFs individuais.

    Retorna os caminhos dos arquivos criados.
    """

    from pypdf import PdfReader, PdfWriter

    reader = PdfReader(
        caminho_pdf,
        strict=False
    )

    arquivos = []

    coluna_nome = detectar_coluna_nome(df)

    total_paginas = len(reader.pages)
    total_linhas = len(df)

    # Segurança contra inconsistências.
    total = min(
        total_paginas,
        total_linhas
    )

    for i in range(total):

        page = reader.pages[i]

        writer = PdfWriter()

        writer.add_page(page)

        nome = str(
            df.iloc[i][coluna_nome]
        )

        nome_limpo = limpar_nome_arquivo(
            nome
        )

        # Evita nome vazio.
        if not nome_limpo:
            nome_limpo = f"participante_{i + 1}"

        caminho = os.path.join(
            pasta_saida,
            f"{nome_limpo}.pdf"
        )

        with open(
            caminho,
            "wb"
        ) as f:

            writer.write(f)

        arquivos.append(caminho)

    return arquivos


# ============================================================
# DOWNLOAD
# ============================================================

@app.route("/download/<arquivo>")
def baixar(arquivo):

    caminho = os.path.join(
        PASTA_DOWNLOAD,
        arquivo
    )

    return send_file(
        caminho,
        as_attachment=True
    )


# ============================================================
# GERAÇÃO DE CERTIFICADOS
# ============================================================

@app.route(
    "/certificados",
    methods=["GET", "POST"]
)
def certificados():

    if "user" not in session:
        return redirect("/login")

    if request.method == "POST":

        fundo = request.files["fundo"]
        planilha = request.files["planilha"]

        texto = request.form["texto"]

        fonte = int(
            request.form["fonte"]
        )

        alinhamento = request.form[
            "alinhamento"
        ]

        posicao_vertical = request.form[
            "posicao_vertical"
        ]

        ajuste_vertical = int(
            request.form["ajuste_vertical"]
        )

        largura_texto_percent = int(
            request.form["largura_texto"]
        )

        # ----------------------------------------------------
        # LEITURA DA PLANILHA
        # ----------------------------------------------------

        df = pd.read_excel(
            planilha
        )

        quantidade = len(df)

        if quantidade == 0:

            return (
                "<h2>Não foi possível gerar os certificados.</h2>"
                "<p>A planilha não possui participantes.</p>"
            ), 400

        # ----------------------------------------------------
        # DIRETÓRIO TEMPORÁRIO
        # ----------------------------------------------------

        pasta_temp = tempfile.mkdtemp()

        caminho_pdf_lote = os.path.join(
            pasta_temp,
            "lote_certificados.pdf"
        )

        caminho_fundo = None

        data = datetime.date.today().strftime(
            "%Y-%m-%d"
        )

        nome_zip = (
            f"certificados_emitte_"
            f"{quantidade}_"
            f"{data}.zip"
        )

        caminho_zip = os.path.join(
            PASTA_DOWNLOAD,
            nome_zip
        )

        try:

            # ------------------------------------------------
            # PREPARAÇÃO DO FUNDO
            # ------------------------------------------------

            imagem = preparar_imagem_fundo(
                fundo
            )

            caminho_fundo = os.path.join(
                pasta_temp,
                "fundo_normalizado.png"
            )

            # Salva uma cópia temporária.
            #
            # A imagem PIL original continua sendo usada
            # pelo ImageReader durante a geração.
            imagem.save(
                caminho_fundo,
                format="PNG"
            )

            # ------------------------------------------------
            # GERAÇÃO DO PDF MULTIPÁGINA
            # ------------------------------------------------

            gerar_pdf_lote(
                imagem,
                df,
                texto,
                fonte,
                alinhamento,
                posicao_vertical,
                ajuste_vertical,
                largura_texto_percent,
                caminho_pdf_lote
            )

            # Depois que o PDF foi criado,
            # a imagem pode ser liberada.
            imagem.close()

            # ------------------------------------------------
            # DIVISÃO DO PDF
            # ------------------------------------------------

            arquivos = dividir_pdf(
                caminho_pdf_lote,
                df,
                pasta_temp
            )

            # ------------------------------------------------
            # CRIAÇÃO DO ZIP
            # ------------------------------------------------
            #
            # ZIP_STORED:
            # os PDFs já são arquivos compactados/estruturados,
            # então não desperdiçamos CPU tentando comprimi-los
            # novamente.
            #
            # zipf.write():
            # adiciona o arquivo sem fazer f.read() para a RAM.
            # ------------------------------------------------

            with zipfile.ZipFile(
                caminho_zip,
                mode="w",
                compression=zipfile.ZIP_STORED
            ) as zipf:

                for arquivo in arquivos:

                    zipf.write(
                        arquivo,
                        arcname=os.path.basename(
                            arquivo
                        )
                    )

            # ------------------------------------------------
            # LIMPEZA DOS PDFs INDIVIDUAIS
            # ------------------------------------------------

            for arquivo in arquivos:

                try:
                    os.remove(arquivo)

                except OSError:
                    pass

            # Remove o PDF multipágina.
            try:

                os.remove(
                    caminho_pdf_lote
                )

            except OSError:
                pass

        except ValueError as erro:

            # Remove ZIP incompleto.
            if os.path.exists(
                caminho_zip
            ):

                try:
                    os.remove(
                        caminho_zip
                    )

                except OSError:
                    pass

            return (
                "<h2>Não foi possível processar "
                "os certificados.</h2>"
                f"<p>{str(erro)}</p>"
            ), 400

        except Exception as erro:

            # Remove ZIP incompleto.
            if os.path.exists(
                caminho_zip
            ):

                try:
                    os.remove(
                        caminho_zip
                    )

                except OSError:
                    pass

            # Log do erro no Render.
            app.logger.exception(
                "Erro durante a geração dos certificados"
            )

            return (
                "<h2>Ocorreu um erro durante "
                "a geração dos certificados.</h2>"
                "<p>Verifique a planilha e a imagem "
                "de fundo e tente novamente.</p>"
            ), 500

        finally:

            # Libera todo o diretório temporário.
            shutil.rmtree(
                pasta_temp,
                ignore_errors=True
            )

        # ----------------------------------------------------
        # SUCESSO
        # ----------------------------------------------------

        return render_template(
            "download.html",
            arquivo=nome_zip
        )

    return render_template(
        "certificados.html"
    )


# ============================================================
# EXECUÇÃO LOCAL
# ============================================================

if __name__ == "__main__":

    port = int(
        os.environ.get(
            "PORT",
            10000
        )
    )

    app.run(
        host="0.0.0.0",
        port=port
    )
