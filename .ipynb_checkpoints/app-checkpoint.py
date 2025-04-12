import streamlit as st
from PIL import Image
from docx import Document
from docx.shared import Inches, RGBColor
from io import BytesIO
import datetime
import math

st.set_page_config(page_title="MED Pictures Generator", layout="wide")
st.title("📸 MED PICTURES Word Document Generator")

layout_options = {
    "2 x 2": (2, 2),
    "3 x 2": (3, 2),
    "4 x 2": (4, 2)
}

project_title = st.text_input("Enter Project Title")
contractor_name = st.text_input("Enter Contractor Name")
layout_choice = st.selectbox("Select Layout (Columns x Rows per Page)", list(layout_options.keys()))
uploaded_images = st.file_uploader("Upload Images", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

cols, rows = layout_options[layout_choice]
images_per_page = cols * rows

if uploaded_images:
    st.subheader("Preview of Selected Layout")
    for i in range(0, len(uploaded_images), cols):
        preview_row = st.columns(cols)
        for j in range(cols):
            if i + j < len(uploaded_images):
                img = Image.open(uploaded_images[i + j])
                preview_row[j].image(img, use_column_width=True)

def create_document():
    doc = Document()

    total_pages = math.ceil(len(uploaded_images) / images_per_page)
    image_index = 0

    for page in range(total_pages):
        if page > 0:
            doc.add_page_break()

        # Header
        paragraph = doc.add_paragraph()
        run = paragraph.add_run("MED PICTURES: ")
        run.font.color.rgb = RGBColor(255, 0, 0)
        run.bold = True
        run.font.size = doc.styles['Normal'].font.size

        run2 = paragraph.add_run(f"{project_title} by {contractor_name}")
        run2.font.color.rgb = RGBColor(0, 0, 0)
        run2.font.size = doc.styles['Normal'].font.size

        # Image Grid
        table = doc.add_table(rows=rows, cols=cols)
        for r in range(rows):
            row_cells = table.rows[r].cells
            for c in range(cols):
                if image_index < len(uploaded_images):
                    image = Image.open(uploaded_images[image_index])
                    image.thumbnail((300, 300))
                    img_stream = BytesIO()
                    image.save(img_stream, format='PNG')
                    img_stream.seek(0)
                    row_cells[c].paragraphs[0].add_run().add_picture(img_stream, width=Inches(2.5))
                    image_index += 1

    doc_stream = BytesIO()
    doc.save(doc_stream)
    doc_stream.seek(0)
    return doc_stream

if st.button("Generate Word Document"):
    if not project_title or not contractor_name or not uploaded_images:
        st.error("Please fill all fields and upload images.")
    else:
        doc_stream = create_document()
        filename = f"MED_PICTURES_{project_title.replace(' ', '_')}.docx"
        st.download_button(label="📥 Download Word Document",
                           data=doc_stream,
                           file_name=filename,
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")