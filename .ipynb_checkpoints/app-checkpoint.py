import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.section import WD_ORIENT
from PIL import Image
from io import BytesIO
import datetime
import math

# Streamlit page config
st.set_page_config(page_title="MED Pictures Generator", layout="wide")
st.title("📸 MED PICTURES Word Document Generator")

# Function to add header
def add_header(doc, project_title, contractor_name):
    section = doc.sections[-1]
    header = section.header
    paragraph = header.paragraphs[0]
    paragraph.clear()
    run_red = paragraph.add_run("MED PICTURES: ")
    run_red.bold = True
    run_red.font.color.rgb = RGBColor(255, 0, 0)
    run_red.font.size = Pt(14)

    run_black = paragraph.add_run(f"{project_title} by {contractor_name}")
    run_black.font.color.rgb = RGBColor(0, 0, 0)
    run_black.font.size = Pt(14)

# Image layout logic
def insert_images(doc, images, layout):
    rows, cols = layout
    images_per_page = rows * cols
    total_pages = math.ceil(len(images) / images_per_page)

    img_idx = 0
    for page in range(total_pages):
        table = doc.add_table(rows=rows, cols=cols)
        table.autofit = True

        for r in range(rows):
            row_cells = table.rows[r].cells
            for c in range(cols):
                if img_idx >= len(images):
                    break
                image = images[img_idx]
                img_idx += 1
                with BytesIO() as output:
                    image.save(output, format="PNG")
                    row_cells[c].paragraphs[0].add_run().add_picture(output, height=Inches(2.5))
        if page < total_pages - 1:
            doc.add_page_break()

# Function to create and return document
def create_document(images, project_title, contractor_name, layout, orientation):
    doc = Document()

    # Page orientation and margin
    section = doc.sections[0]
    if orientation == "Landscape":
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

    # Add header on first section
    add_header(doc, project_title, contractor_name)

    # Add pictures
    insert_images(doc, images, layout)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# Layout options
layout_options = {
    "1 x 2": (1, 2),
    "1 x 3": (1, 3),
    "2 x 2": (2, 2),
    "2 x 3": (2, 3),
    "3 x 2": (3, 2),
    "3 x 3": (3, 3)
}

# --- Main Form Function ---
def picture_form():
    with st.form(key=f"form_{datetime.datetime.now().timestamp()}"):
        project_title = st.text_input("Project Title")
        contractor_name = st.text_input("Contractor Name")
        uploaded_images = st.file_uploader("Upload Images", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

        # View selected images in grid
        if uploaded_images:
            cols = st.columns(3)
            for i, uploaded_file in enumerate(uploaded_images):
                img = Image.open(uploaded_file)
                with cols[i % 3]:
                    st.image(img, caption=uploaded_file.name, use_container_width=True)

        layout_choice = st.selectbox("Select Layout", list(layout_options.keys()))
        orientation = st.radio("Select Page Orientation", ["Portrait", "Landscape"])
        submit_button = st.form_submit_button("Generate Document")

        if submit_button and uploaded_images:
            images = [Image.open(img_file) for img_file in uploaded_images]
            doc_stream = create_document(
                images, project_title, contractor_name,
                layout_options[layout_choice], orientation
            )
            st.download_button("📄 Download Document", doc_stream, file_name="med_pictures.docx")

            # Offer to generate another
            if st.button("Generate Another Document"):
                picture_form()

# --- Run App ---
picture_form()
