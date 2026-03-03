import os
import uuid
import zipfile
import pandas as pd
from flask import Flask, render_template, request, send_file
from PIL import Image, ImageDraw, ImageFont

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
FONT_FOLDER = "fonts"
GENERATED_FOLDER = "generated_certificates"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(FONT_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)

recipient_data = []
background_image = None
selected_font_path = None


def draw_paragraph(draw, text, position, font, max_width, align, line_spacing, color):
    words = text.split()
    lines = []

    while words:
        line = ""
        while words and draw.textlength(line + words[0] + " ", font=font) <= max_width:
            line += words.pop(0) + " "
        lines.append(line.strip())

    y_offset = 0

    for line in lines:
        line_width = draw.textlength(line, font=font)

        if align == "center":
            x = position[0] - line_width / 2
        elif align == "right":
            x = position[0] - line_width
        else:
            x = position[0]

        draw.text((x, position[1] + y_offset),
                  line,
                  fill=color,
                  font=font)

        y_offset += font.size + line_spacing


@app.route("/", methods=["GET", "POST"])
def index():
    global recipient_data, background_image, selected_font_path

    if request.method == "POST":

        if "background" in request.files:
            file = request.files["background"]
            if file.filename:
                background_image = os.path.join(UPLOAD_FOLDER, file.filename)
                file.save(background_image)

        if "excel" in request.files:
            file = request.files["excel"]
            if file.filename:
                df = pd.read_excel(file)
                recipient_data = df.to_dict(orient="records")

        if "fontfile" in request.files:
            file = request.files["fontfile"]
            if file.filename:
                selected_font_path = os.path.join(FONT_FOLDER, file.filename)
                file.save(selected_font_path)

        if "generate" in request.form and background_image:

            text_template = request.form["text"]
            pos_x = int(float(request.form["pos_x"]))
            pos_y = int(float(request.form["pos_y"]))
            font_size = int(request.form["font_size"])
            max_width = int(request.form["max_width"])
            align = request.form["align"]
            line_spacing = int(request.form["line_spacing"])
            color = request.form["color"]

            generated_files = []

            for recipient in recipient_data:
                image = Image.open(background_image).convert("RGB")
                draw = ImageDraw.Draw(image)

                if selected_font_path:
                    font = ImageFont.truetype(selected_font_path, font_size)
                else:
                    font = ImageFont.load_default()

                final_text = text_template
                for key in recipient:
                    final_text = final_text.replace(f"{{{{{key}}}}}", str(recipient[key]))

                filename = f"{uuid.uuid4()}.png"
                save_path = os.path.join(GENERATED_FOLDER, filename)

                draw_paragraph(draw, final_text,
                               (pos_x, pos_y),
                               font,
                               max_width,
                               align,
                               line_spacing,
                               color)

                image.save(save_path)
                generated_files.append(save_path)

            zip_path = os.path.join(GENERATED_FOLDER, "certificados.zip")
            with zipfile.ZipFile(zip_path, "w") as zipf:
                for file in generated_files:
                    zipf.write(file, os.path.basename(file))

            return send_file(zip_path, as_attachment=True)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
