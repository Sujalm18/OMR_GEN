from flask import Flask, render_template, request, send_file, jsonify
import os, zipfile, pandas as pd, re
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Table, TableStyle

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ---------- OMR Configuration ----------
ROLL_X_CM_30 = [7.20, 7.82, 8.44, 9.06, 9.68]
BUBBLE_Y_TOP_CM_30 = [18.8]*5
BUBBLE_SPACING_CM_30 = 0.612
OMR_TEMPLATE_30 = "PHN_Unique_ID_1.jpg"

ROLL_X_CM_50 = [4.6, 5.22, 5.84, 6.46, 7.08]
BUBBLE_Y_TOP_CM_50 = [19.45]*5
BUBBLE_SPACING_CM_50 = 0.622
OMR_TEMPLATE_50 = "PHN_Unique_ID_2.jpg"

# ---------- Helpers ----------
def normalize_col_name(s):
    return re.sub(r'[^a-z0-9]', '', str(s).lower().strip()) if s else ""

def find_column(df_cols_norm, aliases):
    for orig_col, norm in df_cols_norm.items():
        for a in aliases:
            if norm == a: return orig_col
    for orig_col, norm in df_cols_norm.items():
        for a in aliases:
            if a in norm: return orig_col
    return None

def safe_filename(s):
    s = str(s).strip()
    s = re.sub(r'[\\/*?:"<>|]', '_', s)
    s = re.sub(r'\s+', '_', s)
    return s[:200]

def format_roll_value(v):
    if pd.isna(v): return "00000"
    try: return str(int(float(v))).zfill(5)
    except: return str(v).zfill(5)

def fill_roll_bubbles(c, roll_no, ROLL_X_CM, BUBBLE_Y_TOP_CM, BUBBLE_SPACING_CM):
    roll_no = str(roll_no).zfill(5)
    for i, digit_char in enumerate(roll_no):
        digit = int(digit_char)
        x = ROLL_X_CM[i] * cm
        y = BUBBLE_Y_TOP_CM[i]*cm - digit * BUBBLE_SPACING_CM * cm
        c.setFillColor(colors.black)
        c.circle(x, y, 0.23*cm, stroke=0, fill=1)
    c.setFillColor(colors.black)

def draw_roll_number_text(c, roll_no, ROLL_X_CM, BUBBLE_Y_TOP_CM):
    roll_no = str(roll_no).zfill(5)
    text_y = (BUBBLE_Y_TOP_CM[0]*cm) + 0.40*cm
    c.setFont("Helvetica-Bold", 13)
    c.setFillColor(colors.black)
    for i, digit_char in enumerate(roll_no):
        x = ROLL_X_CM[i]*cm
        c.drawCentredString(x, text_y, digit_char)

def draw_single_omr(c, row, class_numeric):
    # Decide template based on class
    if class_numeric in [1,2,3]:
        ROLL_X_CM = ROLL_X_CM_30
        BUBBLE_Y_TOP_CM = BUBBLE_Y_TOP_CM_30
        BUBBLE_SPACING_CM = BUBBLE_SPACING_CM_30
        omr_template = OMR_TEMPLATE_30
    else:
        ROLL_X_CM = ROLL_X_CM_50
        BUBBLE_Y_TOP_CM = BUBBLE_Y_TOP_CM_50
        BUBBLE_SPACING_CM = BUBBLE_SPACING_CM_50
        omr_template = OMR_TEMPLATE_50

    if not os.path.exists(omr_template):
        return

    omr_img = ImageReader(omr_template)
    width, height = A4
    c.drawImage(omr_img, 0, 0, width=width, height=height, preserveAspectRatio=True)

    roll_no = format_roll_value(row.get('Roll. no.', '00000'))
    fill_roll_bubbles(c, roll_no, ROLL_X_CM, BUBBLE_Y_TOP_CM, BUBBLE_SPACING_CM)
    draw_roll_number_text(c, roll_no, ROLL_X_CM, BUBBLE_Y_TOP_CM)

    data = [
        [f"Student Name: {row.get('Name of the student', ' ')}"],
        [f"School: {row.get('School Name', ' ')}"],
        [f"Class: {row.get('Class', ' ')}    Division: {row.get('Division', ' ')}    Roll No.:"],
        ["Question Paper Set: _____________"]
    ]
    table_width = width*0.7
    table = Table(data, colWidths=[table_width])
    table.setStyle(TableStyle([
        ("BOX", (0,0), (-1,-1), 0.8, colors.black),
        ("INNERGRID",(0,0),(-1,-1),0.5, colors.black),
        ("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),11),
        ("ALIGN",(0,0),(-1,-1),"LEFT"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",(0,0),(-1,-1),5),
        ("BOTTOMPADDING",(0,0),(-1,-1),5),
    ]))
    w,h = table.wrap(0,0)
    x = (width-w)/2
    y = height-4.5*cm - h
    table.drawOn(c,x,y)

def generate_combined_pdf_for_sheet(df, output_path):
    c = canvas.Canvas(output_path, pagesize=A4)
    for idx,row in df.iterrows():
        try:
            class_numeric = int(str(row.get('Class','4')).strip())
        except:
            class_numeric = 4
        draw_single_omr(c,row,class_numeric)
        c.showPage()
    c.save()

def create_zip_of_pdfs(pdf_dir, zip_filename):
    with zipfile.ZipFile(zip_filename,'w',zipfile.ZIP_DEFLATED) as zipf:
        for foldername, _, filenames in os.walk(pdf_dir):
            for filename in filenames:
                if filename.endswith(".pdf"):
                    file_path = os.path.join(foldername,filename)
                    arcname = os.path.relpath(file_path,pdf_dir)
                    zipf.write(file_path, arcname)

# ---------- Routes ----------
@app.route('/')
def upload_page():
    return render_template('uploads.html')

@app.route('/generate', methods=['POST'])
def generate_omrs():
    file = request.files['excel_file']
    if not file:
        return "No file uploaded",400
    filepath = os.path.join(UPLOAD_FOLDER,file.filename)
    file.save(filepath)

    xls = pd.ExcelFile(filepath)
    output_zip = BytesIO()
    temp_pdf_folder = os.path.join(UPLOAD_FOLDER,"temp_pdfs")
    os.makedirs(temp_pdf_folder, exist_ok=True)

    with zipfile.ZipFile(output_zip,'w',zipfile.ZIP_DEFLATED) as zipf:
        for idx, sheet_name in enumerate(xls.sheet_names):
            df = pd.read_excel(xls, sheet_name=sheet_name, dtype=object)
            pdf_filename = f"{safe_filename(sheet_name)}.pdf"
            pdf_path = os.path.join(temp_pdf_folder,pdf_filename)
            generate_combined_pdf_for_sheet(df,pdf_path)
            zipf.write(pdf_path,pdf_filename)
            os.remove(pdf_path)

    output_zip.seek(0)
    return send_file(output_zip, download_name="OMR_Sheets.zip", as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
