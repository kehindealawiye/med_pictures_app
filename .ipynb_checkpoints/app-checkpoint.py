import streamlit as st
from PIL import Image
from docx import Document
from docx.shared import Inches, RGBColor
from io import BytesIO
import os
import math

# Streamlit page configuration
st.set_page_config(page_title="MED Pictures Generator", layout="wide")
st.title("📸 MED PICTURES Word Document Generator")

def add_header(doc, project_title, contractor_name):
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("MED PICTURES: ")
    run.font.color.rgb = RGBColor(255, 0, 0)  # Red text for 'MED PICTURES'
    run.bold = True
    run = paragraph.add_run(f"{project_title} by {contractor_name}")
    run.font.color.rgb = RGBColor(0, 0, 0)  # Black text for rest
    paragraph.alignment = 0  # Left align

def create_document(project_title, contractor_name, images, layout, orientation):
    doc = Document()

    # Set orientation
    section = doc.sections[0]
    if orientation == 'Landscape':
        section.orientation = 1
        section.page_width, section.page_height = section.page_height, section.page_width

    # Set margins
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)

    # Determine layout grid
    layout_map = {
        "2 x 2": (2, 2),
        "3 x 2": (3, 2),
        "3 x 3": (3, 3),
        "4 x 2": (4, 2)
    }
    cols, rows = layout_map[layout]
    images_per_page = cols * rows

    for i in range(0, len(images), images_per_page):
        if i != 0:
            doc.add_page_break()

        add_header(doc, project_title, contractor_name)
        table = doc.add_table(rows=rows, cols=cols)
        table.autofit = True

        chunk = images[i:i+images_per_page]

        for idx, image_file in enumerate(chunk):
            row = idx // cols
            col = idx % cols
            cell = table.cell(row, col)

            image = Image.open(image_file)
            aspect_ratio = image.width / image.height
            target_height = 3  # Inches
            target_width = target_height * aspect_ratio

            img_stream = BytesIO()
            image.save(img_stream, format='PNG')
            img_stream.seek(0)

            cell.paragraphs[0].add_run().add_picture(img_stream, height=Inches(target_height))

    file_stream = BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# User Inputs
with st.form("med_pictures_form"):
    project_title = st.text_input("Project Title")
    contractor_name = st.text_input("Contractor Name")
    uploaded_images = st.file_uploader("Upload Images", accept_multiple_files=True, type=["png", "jpg", "jpeg"])
    layout_option = st.selectbox("Select Layout", ["2 x 2", "3 x 2", "3 x 3", "4 x 2"])
    orientation = st.selectbox("Page Orientation", ["Portrait", "Landscape"])
    submitted = st.form_submit_button("Generate Word Document")

if submitted:
    if not project_title or not contractor_name or not uploaded_images:
        st.error("Please fill all fields and upload at least one image.")
    else:
        doc_stream = create_document(project_title, contractor_name, uploaded_images, layout_option, orientation)
        filename = f"MED_PICTURES_{project_title}_by_{contractor_name}.docx".replace(" ", "_")
        st.success("Document generated successfully!")
        st.download_button("Download Document", data=doc_stream, file_name=filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
