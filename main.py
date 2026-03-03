import os
import uuid
import pandas as pd
from flask import Flask, render_template, request, send_from_directory
from PIL import Image, ImageDraw, ImageFont

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
GENERATED_FOLDER = "generated_certificates"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)

recipient_data = []
background_image = None

@app.route("/", methods=["GET", "POST"])
def index():
    global recipient_data
    global background_image

    if request.method == "POST":

        # Upload fundo
        if "background" in request.files:
            file = request.files["background"]
            if file.filename != "":
                background_image = os.path.join(UPLOAD_FOLDER, file.filename)
                file.save(background_image)

        # Upload Excel
        if "excel" in request.files:
            file = request.files["excel"]
            if file.filename != "":
                df = pd.read_excel(file)
                recipient_data = df.to_dict(orient="records")

        # Gerar certificados
        if "generate" in request.form and background_image:
            text_template = request.form["text"]
            pos_x = int(request.form["pos_x"])
            pos_y = int(request.form["pos_y"])
            font_size = int(request.form["font_size"])

            for recipient in recipient_data:
                image = Image.open(background_image).convert("RGB")
                draw = ImageDraw.Draw(image)

                try:
                    font = ImageFont.truetype("arial.ttf", font_size)
                except:
                    font = ImageFont.load_default()

                final_text = text_template
                for key in recipient:
                    final_text = final_text.replace(f"{{{{{key}}}}}", str(recipient[key]))

                filename = f"{uuid.uuid4()}.png"
                save_path = os.path.join(GENERATED_FOLDER, filename)

                draw.text((pos_x, pos_y), final_text, fill="black", font=font)
                image.save(save_path)

    return render_template("index.html")

@app.route("/downloads/<filename>")
def download_file(filename):
    return send_from_directory(GENERATED_FOLDER, filename)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
