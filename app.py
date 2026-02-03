import os
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify
from fpdf import FPDF
import tempfile
import uuid
import base64
import qrcode
import barcode
from barcode.writer import ImageWriter

app = Flask(__name__)

class PDF(FPDF):
    def __init__(self, box_width, box_height, font_size, border_width, columns, header_text, font_family, font_style):
        super().__init__()
        self.box_width, self.box_height = box_width, box_height
        self.font_size, self.border_width = font_size, border_width
        self.columns, self.header_text = columns, header_text
        self.font_family, self.font_style = font_family, font_style

    def header(self):
        self.set_font(self.font_family, style=self.font_style, size=14)
        self.set_y(10)
        self.cell(0, 10, self.header_text, ln=True, align="C")

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f"Designed by Hirusha | Powered by FCT StickerGen Pro | Page {self.page_no()}", align="C")

def generate_pdf_logic(data_rows, output_path, config):
    pdf = PDF(config['width'], config['height'], config['font_size'], 
              config['border'], config['cols'], config['header'], 
              'Arial', "")
    
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()
    pdf.set_line_width(config['border'])

    x_margin, y_start, spacing = 10, 25, 1
    y, counter = y_start, 1
    
    for i in range(0, len(data_rows), config['cols']):
        if y + config['height'] > pdf.h - 25:
            pdf.add_page()
            y = y_start
        x = x_margin
        for col_index in range(config['cols']):
            if i + col_index < len(data_rows):
                lines = data_rows[i + col_index]
                pdf.set_xy(x, y)
                pdf.cell(w=config['width'], h=config['height'], txt="", border=1)

                mode = config.get('mode')
                
                if mode == 'barcode':
                    temp_img = os.path.join(tempfile.gettempdir(), f"{uuid.uuid4().hex}.png")
                    try:
                        code_class = barcode.get('code128', str(lines[0]), writer=ImageWriter())
                        code_class.save(temp_img.replace('.png', ''), options={"write_text": False})
                        barcode_h = config['height'] * 0.45
                        pdf.image(temp_img, x + 2, y + 2, w=config['width'] - 4, h=barcode_h)
                        current_y = y + barcode_h + 3
                        draw_x = x + 2
                        text_width = config['width'] - 4
                        align = "C"
                    except: pass
                    finally:
                        if os.path.exists(temp_img): os.remove(temp_img)
                
                elif mode == 'qr':
                    temp_img = os.path.join(tempfile.gettempdir(), f"{uuid.uuid4().hex}.png")
                    img_size = config['height'] - 4
                    qr = qrcode.make(lines[0])
                    qr.save(temp_img)
                    pdf.image(temp_img, x + 2, y + 2, h=img_size)
                    line_x = x + img_size + 3
                    pdf.set_draw_color(200, 200, 200)
                    pdf.line(line_x, y + 2, line_x, y + config['height'] - 2)
                    pdf.set_draw_color(0, 0, 0)
                    draw_x = x + img_size + 5
                    text_width = config['width'] - img_size - 7
                    current_y = y + 4
                    align = "L"
                    if os.path.exists(temp_img): os.remove(temp_img)
                else: 
                    draw_x = x + 2
                    text_width = config['width'] - 4
                    current_y = y + (config['height'] / 4)
                    align = "C"

                line1_fsize = config['font_size']
                sub_fsize = line1_fsize * 0.7 if line1_fsize > 8 else 6
                for idx, text in enumerate(lines):
                    is_main = (idx == 0)
                    pdf.set_font("Arial", "B" if is_main else "", line1_fsize if is_main else sub_fsize)
                    while pdf.get_string_width(str(text)) > text_width and pdf.font_size_pt > 4:
                        pdf.set_font_size(pdf.font_size_pt - 0.5)
                    pdf.set_xy(draw_x, current_y)
                    pdf.cell(w=text_width, h=line1_fsize*0.35, txt=str(text), border=0, align=align)
                    current_y += (line1_fsize * 0.4)

                if config.get('show_corners'):
                    pdf.set_font("Arial", "B", 5)
                    pdf.set_text_color(180, 180, 180)
                    pdf.set_xy(x + 1, y + 0.5)
                    pdf.cell(5, 3, str(counter).zfill(2))
                    pdf.set_text_color(0, 0, 0)
                x += config['width'] + spacing
                counter += 1
        y += config['height'] + spacing
    pdf.output(output_path)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/get_sheets', methods=['POST'])
def get_sheets():
    file = request.files.get('file')
    if not file: return jsonify([])
    file.seek(0)
    ext = os.path.splitext(file.filename)[1].lower()
    if ext in ['.xlsx', '.xls']:
        xl = pd.ExcelFile(file)
        return jsonify(xl.sheet_names)
    return jsonify(["Default"])

@app.route('/get_columns', methods=['POST'])
def get_columns():
    file = request.files.get('file')
    sheet = request.form.get('sheet_name')
    if not file: return jsonify([])
    file.seek(0)
    ext = os.path.splitext(file.filename)[1].lower()
    try:
        df = pd.read_excel(file, sheet_name=sheet) if ext in ['.xlsx', '.xls'] else pd.read_csv(file)
        return jsonify(df.columns.tolist())
    except: return jsonify([])

@app.route('/generate', methods=['POST'])
def generate():
    file = request.files.get('file')
    sheet = request.form.get('sheet_selector')
    col1 = request.form.get('col_selector_1')
    col2 = request.form.get('col_selector_2')
    col3 = request.form.get('col_selector_3')
    action = request.form.get('action')
    if not file or not col1: return jsonify({"error": "Missing selection"}), 400
    file.seek(0)
    ext = os.path.splitext(file.filename)[1].lower()
    df = pd.read_excel(file, sheet_name=sheet) if ext in ['.xlsx', '.xls'] else pd.read_csv(file)
    data_rows = []
    for _, row in df.iterrows():
        entry = [str(row[col1])]
        if col2 and col2 != "None": entry.append(str(row[col2]))
        if col3 and col3 != "None": entry.append(str(row[col3]))
        data_rows.append(entry)
    config = {
        'width': float(request.form.get('width', 90)),
        'height': float(request.form.get('height', 19)),
        'font_size': float(request.form.get('font_size', 10)),
        'border': 0.2,
        'cols': int(request.form.get('cols', 2)),
        'header': request.form.get('header', "Label Sheet"),
        'mode': request.form.get('mode', 'text'),
        'show_corners': 'show_corners' in request.form
    }
    tmp_path = os.path.join(tempfile.gettempdir(), f"{uuid.uuid4().hex}.pdf")
    generate_pdf_logic(data_rows, tmp_path, config)
    if action == 'download':
        # FIXED: Added mimetype for PDF streams
        return send_file(tmp_path, as_attachment=True, download_name=f"{config['header']}.pdf", mimetype='application/pdf')
    else:
        with open(tmp_path, "rb") as f:
            pdf_base64 = base64.b64encode(f.read()).decode('utf-8')
        if os.path.exists(tmp_path): os.remove(tmp_path)
        return jsonify({'pdf': pdf_base64})

if __name__ == '__main__':
    app.run(debug=True, port=5000)