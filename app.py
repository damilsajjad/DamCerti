from flask import Flask, render_template, request, send_file
import os
import openpyxl
from PIL import Image, ImageDraw, ImageFont
import zipfile
import io
from dotenv import load_dotenv
from supabase import create_client

# Load environment variables from .env file
load_dotenv()

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


@app.route('/')
def home():
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def generate():
    template_image = request.files['template_image']
    excel_file = request.files['excel_file']
    x_percent = float(request.form['x_percent'])
    y_percent = float(request.form['y_percent'])
    font_file = request.form['font_file']
    font_size = int(float(request.form['font_size']))
    text_color = request.form['text_color']  # e.g. "#ff0000"

    # Save uploads
    image_path = os.path.join(UPLOAD_FOLDER, template_image.filename)
    excel_path = os.path.join(UPLOAD_FOLDER, excel_file.filename)
    template_image.save(image_path)
    excel_file.save(excel_path)

    # Read names from Excel
    wb = openpyxl.load_workbook(excel_path)
    sheet = wb.active
    headers = [cell.value for cell in sheet[1]]
    name_col_index = headers.index('Name')

    names = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[name_col_index]:
            names.append(str(row[name_col_index]))

    # Load font (fallback to Arial if the chosen font file isn't found)
    try:
        font = ImageFont.truetype(font_file, font_size)
    except Exception:
        font = ImageFont.truetype("arial.ttf", font_size)

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
        for name in names:
            img = Image.open(image_path).convert("RGB")
            draw = ImageDraw.Draw(img)

            x = (x_percent / 100) * img.width
            y = (y_percent / 100) * img.height

            bbox = draw.textbbox((0, 0), name, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]

            # x,y is the CENTER (matches the draggable preview), shift to top-left for drawing
            draw_x = x - text_width / 2
            draw_y = y - text_height / 2

            draw.text((draw_x, draw_y), name, fill=text_color, font=font)

            pdf_buffer = io.BytesIO()
            img.save(pdf_buffer, format="PDF")
            pdf_buffer.seek(0)
            zip_file.writestr(f"Certificate_{name}.pdf", pdf_buffer.read())

    zip_buffer.seek(0)
    return send_file(
        zip_buffer,
        mimetype='application/zip',
        as_attachment=True,
        download_name='certificates.zip'
    )


if __name__ == '__main__':
    app.run(debug=True)
