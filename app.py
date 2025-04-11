import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from PIL import Image
import io
from datetime import datetime
import os

st.set_page_config(layout="centered")
st.title("📸 MED PICTURES DOCUMENT GENERATOR")

# --- Sidebar Inputs ---
st.sidebar.header("📌 Settings")
project_title = st.sidebar.text_input("Project Title")
contractor_name = st.sidebar.text_input("Contractor Name")
layout_option = st.sidebar.selectbox("Select Layout", ["2 x 2", "3 x 2", "3 x 3"])
orientation = st.sidebar.radio("Page Orientation", ["Portrait", "Landscape"])
image_width_inches = st.sidebar.slider("Image Width (inches)", 2.0, 5.0, 3.0)

uploaded_images = st.file_uploader("Upload Project Images", accept_multiple_files=True, type=["png", "jpg", "jpeg"])

if st.button("Generate Document") and uploaded_images:
    cols = int(layout_option.split(" x ")[0])
    rows = int(layout_option.split(" x ")[1])
    max_per_page = cols * rows
    img_width = Inches(image_width_inches)

    # Document setup
    doc = Document()
    if orientation == "Landscape":
        section = doc.sections[-1]
        section.orientation = 1  # Landscape
        new_width, new_height = section.page_height, section.page_width
        section.page_width, section.page_height = new_width, new_height

    # Title setup
    def add_title():
        title_para = doc.add_paragraph()
        run1 = title_para.add_run("MED PICTURES: ")
        run1.font.color.rgb = RGBColor(255, 0, 0)
        run1.bold = True
        run1.font.size = Pt(16)

        run2 = title_para.add_run(f"{project_title} by {contractor_name}")
        run2.font.color.rgb = RGBColor(0, 0, 0)
        run2.font.size = Pt(16)
        title_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

        # Add datetime
        dt_para = doc.add_paragraph(datetime.now().strftime("%d-%b-%Y %I:%M %p"))
        dt_para.runs[0].font.size = Pt(10)

    add_title()

    # Add images in grid
    for i, image_file in enumerate(uploaded_images):
        if i % max_per_page == 0 and i != 0:
            doc.add_page_break()
            add_title()

        if i % cols == 0:
            table = doc.add_table(rows=1, cols=cols)
            table.autofit = True
            row_cells = table.rows[0].cells

        img = Image.open(image_file)
        img_io = io.BytesIO()
        img.save(img_io, format='PNG')
        img_io.seek(0)

        # Auto-scale height
        aspect_ratio = img.height / img.width
        height = Inches(image_width_inches * aspect_ratio)

        row_cells[i % cols].paragraphs[0].add_run().add_picture(img_io, width=img_width, height=height)

    # Save document
    safe_title = f"{project_title}".replace('/', '_').replace(':', '-')
    safe_contractor = f"{contractor_name}".replace('/', '_').replace(':', '-')
    filename = f"MED_PICTURES_{safe_title} by {safe_contractor}.docx"
    filepath = os.path.join("/mnt/data", filename)
    doc.save(filepath)

    st.success("✅ Document generated successfully!")
    st.download_button("📥 Download Word Document", data=open(filepath, "rb"), file_name=filename)
