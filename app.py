import streamlit as st
import os
import re
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Table, TableStyle
import tempfile
import zipfile

# Center the company logo
st.markdown(
    """
    <div style='text-align: center;'>
        <img src='logo.webp' width='200'>
    </div>
    """,
    unsafe_allow_html=True
)

st.title("PHN Scholar Exam OMR Generation Software")

ROLL_X_CM_30 = [7.20, 7.82, 8.44, 9.06, 9.68]
BUBBLE_Y_TOP_CM_30 = [18.8] * 5
BUBBLE_SPACING_CM_30 = 0.612
BUBBLE_RADIUS_CM = 0.23
OMR_TEMPLATE_30 = "PHN_Unique_ID_1.jpg"

ROLL_X_CM_50 = [4.6, 5.22, 5.84, 6.46, 7.08]
BUBBLE_Y_TOP_CM_50 = [19.45] * 5
BUBBLE_SPACING_CM_50 = 0.622
OMR_TEMPLATE_50 = "PHN_Unique_ID_2.jpg"

def normalize_col_name(s):
    return re.sub(r'[^a-z0-9]', '', str(s).lower().strip()) if s else ""

def find_column(df_cols_norm, aliases):
    for orig_col, norm in df_cols_norm.items():
        for a in aliases:
            if norm == a:
                return orig_col
    for orig_col, norm in df_cols_norm.items():
        for a in aliases:
            if a in norm:
                return orig_col
    return None

def safe_filename(s):
    s = str(s).strip()
    s = re.sub(r'[\\/*?:"<>|]', '_', s)
    s = re.sub(r'\s+', '_', s)
    return s[:200]

def format_roll_value(v):
    if pd.isna(v):
        return "00000"
    try:
        return str(int(float(v))).zfill(5)
    except:
        return str(v).zfill(5)

def fill_roll_bubbles(c, roll_no, ROLL_X_CM, BUBBLE_Y_TOP_CM, BUBBLE_SPACING_CM):
    roll_no = str(roll_no).zfill(5)
    for i, digit_char in enumerate(roll_no):
        digit = int(digit_char)
        x = ROLL_X_CM[i] * cm
        y = BUBBLE_Y_TOP_CM[i] * cm - digit * BUBBLE_SPACING_CM * cm
        c.setFillColor(colors.black)
        c.circle(x, y, BUBBLE_RADIUS_CM * cm, stroke=0, fill=1)
    c.setFillColor(colors.black)

def draw_roll_number_text(c, roll_no, ROLL_X_CM, BUBBLE_Y_TOP_CM):
    roll_no = str(roll_no).zfill(5)
    text_y = (BUBBLE_Y_TOP_CM[0] * cm) + (0.40 * cm)
    c.setFont("Helvetica-Bold", 13)
    c.setFillColor(colors.black)
    for i, digit_char in enumerate(roll_no):
        x = ROLL_X_CM[i] * cm
        c.drawCentredString(x, text_y, digit_char)

def create_zip_of_pdfs(pdf_dir, zip_filename):
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for foldername, _, filenames in os.walk(pdf_dir):
            for filename in filenames:
                if filename.endswith(".pdf"):
                    file_path = os.path.join(foldername, filename)
                    arcname = os.path.relpath(file_path, pdf_dir)
                    zipf.write(file_path, arcname)

uploaded_file = st.file_uploader("Upload Excel file", type=["xls", "xlsx"])
if uploaded_file:
    with st.spinner("Processing..."):
        with tempfile.TemporaryDirectory() as tmpdir:
            xls_path = os.path.join(tmpdir, uploaded_file.name)
            with open(xls_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            output_dir = os.path.join(tmpdir, "Generated_OMRs")
            os.makedirs(output_dir, exist_ok=True)

            try:
                xls = pd.ExcelFile(xls_path)
            except Exception as e:
                st.error(f"Cannot read Excel file: {e}")
                st.stop()

            for sheet_name in xls.sheet_names:
                try:
                    df = pd.read_excel(xls_path, sheet_name=sheet_name, dtype=object)
                except Exception as e:
                    st.warning(f"Failed to read sheet: {sheet_name} ({e})")
                    continue

                df_cols_norm = {orig: normalize_col_name(orig) for orig in df.columns}
                aliases = {
                    "school_name": ["schoolname", "school"],
                    "class": ["class"],
                    "division": ["division"],
                    "roll_no": ["rollno", "rollnumber", "roll_no", "uniqueid", "phnuniqueid"],
                    "student_name": ["nameofthestudent", "name", "studentname"],
                }
                col_map = {canon: find_column(df_cols_norm, al_list) for canon, al_list in aliases.items()}
                grouped = df.groupby(col_map["class"]) if col_map["class"] else None
                if grouped is None:
                    continue

                for class_val, group in grouped:
                    class_numeric = None
                    try:
                        class_numeric = int(str(class_val).strip())
                    except:
                        roman_map = {"I": 1, "II": 2, "III": 3}
                        class_upper = str(class_val).strip().upper()
                        class_numeric = roman_map.get(class_upper, None)
                    if class_numeric is None:
                        class_numeric = 4

                    if class_numeric in [1, 2, 3]:
                        ROLL_X_CM = ROLL_X_CM_30
                        BUBBLE_Y_TOP_CM = BUBBLE_Y_TOP_CM_30
                        BUBBLE_SPACING_CM = BUBBLE_SPACING_CM_30
                        omr_template = OMR_TEMPLATE_30
                    else:
                        ROLL_X_CM = ROLL_X_CM_50
                        BUBBLE_Y_TOP_CM = BUBBLE_Y_TOP_CM_50
                        BUBBLE_SPACING_CM = BUBBLE_SPACING_CM_50
                        omr_template = OMR_TEMPLATE_50

                    omr_template_path = omr_template
                    if not os.path.exists(omr_template_path):
                        st.warning(f"OMR template not found: {omr_template_path}")
                        continue

                    omr_img = ImageReader(omr_template_path)
                    pdf_filename = os.path.join(output_dir, f"{safe_filename(sheet_name)}_Class_{safe_filename(str(class_val))}.pdf")
                    c = canvas.Canvas(pdf_filename, pagesize=A4)
                    width, height = A4

                    for idx, row in group.iterrows():
                        student_name = row[col_map["student_name"]] if col_map["student_name"] else ""
                        school_name = row[col_map["school_name"]] if col_map["school_name"] else ""
                        class_name = row[col_map["class"]] if col_map["class"] else ""
                        division = row[col_map["division"]] if col_map["division"] else ""
                        roll_no_raw = row[col_map["roll_no"]] if col_map["roll_no"] else ""
                        roll_no = format_roll_value(roll_no_raw)

                        c.drawImage(omr_img, 0, 0, width=width, height=height, preserveAspectRatio=True)
                        fill_roll_bubbles(c, roll_no, ROLL_X_CM, BUBBLE_Y_TOP_CM, BUBBLE_SPACING_CM)
                        draw_roll_number_text(c, roll_no, ROLL_X_CM, BUBBLE_Y_TOP_CM)

                        data = [
                            [f"Student Name: {student_name or ' '}"],
                            [f"School: {school_name or ' '}"],
                            [f"Class: {class_name or ' '}    Division: {division or ' '}    Roll No.:"],
                            ["Question Paper Set: _____________"],
                        ]
                        table_width = width * 0.7
                        table = Table(data, colWidths=[table_width])
                        table.setStyle(
                            TableStyle(
                                [
                                    ("BOX", (0, 0), (-1, -1), 0.8, colors.black),
                                    ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.black),
                                    ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
                                    ("FONTSIZE", (0, 0), (-1, -1), 11),
                                    ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                                ]
                            )
                        )
                        w, h = table.wrap(0, 0)
                        x = (width - w) / 2
                        y = height - 4.5 * cm - h
                        table.drawOn(c, x, y)
                        c.showPage()
                    c.save()

            zip_path = os.path.join(tmpdir, "OMR_Classwise_PDFs.zip")
            create_zip_of_pdfs(output_dir, zip_path)
            with open(zip_path, "rb") as f:
                st.success("Done! Download the ZIP below.")
                st.download_button(
                    label="Download ZIP",
                    data=f.read(),
                    file_name="OMR_Classwise_PDFs.zip",
                    mime="application/zip"
                )

