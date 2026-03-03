import os
import uuid
import zipfile
import pandas as pd
from flask import Flask, render_template, request, send_file, send_from_directory
from PIL import Image, ImageDraw, ImageFont

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
GENERATED_FOLDER = "generated_certificates"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)

recipient_data = []
background_image = None


def draw_multiline_text(draw, text, position, font, max_width, align="center"):
    lines = []
    words = text.split()
    while words:
        line = ""
        while words and draw.textlength(line + words[0], font=font) <= max_width:
            line += words.pop(0) + " "
        lines.append(line)

    y_offset = 0
    for line in lines:
        w = draw.textlength(line, font=font)
        x = position[0] - w // 2 if align == "center" else position[0]
        draw.text((x, position[1] + y_offset), line, fill="black", font=font)
        y_offset += font.size + 10


@app.route("/", methods=["GET", "POST"])
def index():
    global recipient_data
    global background_image

    if request.method == "POST":

        if "background" in request.files:
            file = request.files["background"]
            if file.filename != "":
                background_image = os.path.join(UPLOAD_FOLDER, file.filename)
                file.save(background_image)

        if "excel" in request.files:
            file = request.files["excel"]
            if file.filename != "":
                df = pd.read_excel(file)
                recipient_data = df.to_dict(orient="records")

        if "generate" in request.form and background_image:
            text_template = request.form["text"]
            pos_x = int(request.form["pos_x"])
            pos_y = int(request.form["pos_y"])
            font_size = int(request.form["font_size"])

            generated_files = []

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

                draw_multiline_text(draw, final_text, (pos_x, pos_y), font, 800)

                image.save(save_path)
                generated_files.append(save_path)

            zip_path = os.path.join(GENERATED_FOLDER, "certificados.zip")
            with zipfile.ZipFile(zip_path, "w") as zipf:
                for file in generated_files:
                    zipf.write(file, os.path.basename(file))

            return send_file(zip_path, as_attachment=True)

    return render_template("index.html", background_image=background_image)


@app.route("/certificados")
def view_certificates():
    files = os.listdir(GENERATED_FOLDER)
    return render_template("certificados.html", files=files)


@app.route("/download/<filename>")
def download_file(filename):
    return send_from_directory(GENERATED_FOLDER, filename)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
