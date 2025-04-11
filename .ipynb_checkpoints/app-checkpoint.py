import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENT
from PIL import Image
import os
import io
from datetime import datetime

st.set_page_config(page_title="MED Pictures Generator", layout="wide")
st.title("📸 MED PICTURES Word Document Generator")

# Form Inputs
with st.form("input_form"):
    col1, col2 = st.columns(2)
    with col1:
        project_title = st.text_input("Project Title", "")
        contractor_name = st.text_input("Contractor Name", "")
        image_files = st.file_uploader("Upload Project Images", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
    with col2:
        image_width_in = st.slider("Select Image Width (inches)", 2.0, 6.0, 3.0, 0.5)
        layout_option = st.selectbox("Choose Image Layout", ["2 x 2", "3 x 2", "4 x 2"])
        orientation = st.radio("Page Orientation", ["Landscape", "Portrait"])

    generate_button = st.form_submit_button("Generate Document")

if generate_button and project_title and contractor_name and image_files:
    doc = Document()

    # Orientation setup
    section = doc.sections[0]
    if orientation == "Landscape":
        section.orientation = WD_ORIENT.LANDSCAPE
        new_width, new_height = section.page_height, section.page_width
    else:
        section.orientation = WD_ORIENT.PORTRAIT
        new_width, new_height = section.page_width, section.page_height
    section.page_width = new_width
    section.page_height = new_height

    # Title
    full_title = f"MED PICTURES: "
    subtitle = f"{project_title} by {contractor_name}"
    p = doc.add_paragraph()
    p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    run = p.add_run(full_title)
    run.font.color.rgb = docx.shared.RGBColor(255, 0, 0)
    run.font.size = Pt(16)
    run.bold = True
    run = p.add_run(subtitle)
    run.font.color.rgb = docx.shared.RGBColor(0, 0, 0)
    run.font.size = Pt(16)

    # Date and time
    date_str = datetime.now().strftime("Generated on %Y-%m-%d at %H:%M")
    doc.add_paragraph(date_str)

    # Layout config
    layout_map = {"2 x 2": (2, 2), "3 x 2": (3, 2), "4 x 2": (4, 2)}
    cols_per_row, rows_per_page = layout_map[layout_option]
    images_per_page = cols_per_row * rows_per_page

    def resize_to_height(img, target_height_in):
        dpi = 96  # default screen DPI
        target_height_px = int(target_height_in * dpi)
        aspect_ratio = img.width / img.height
        new_width_px = int(target_height_px * aspect_ratio)
        return img.resize((new_width_px, target_height_px))

    # Insert images
    for i, image_file in enumerate(image_files):
        if i % images_per_page == 0 and i != 0:
            doc.add_page_break()
            # Repeat project header on new page
            p = doc.add_paragraph()
            p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            run = p.add_run(full_title)
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.size = Pt(16)
            run.bold = True
            run = p.add_run(subtitle)
            run.font.color.rgb = docx.shared.RGBColor(0, 0, 0)
            run.font.size = Pt(16)
            doc.add_paragraph(date_str)

        img = Image.open(image_file)
        img = resize_to_height(img, image_width_in)
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        buf.seek(0)

        doc.add_picture(buf, width=Inches(image_width_in))
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Save to buffer for download
    filename = f"MED_PICTURES_{project_title}_{contractor_name}.docx".replace(" ", "_")
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)

    st.success("✅ Document generated successfully!")
    st.download_button("📥 Download Word Document", data=doc_io, file_name=filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
