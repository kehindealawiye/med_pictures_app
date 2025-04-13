import streamlit as st
from docx import Document
from docx.shared import Inches, RGBColor, Pt
from PIL import Image
import io
import math
import datetime

st.set_page_config(page_title="MED Pictures Generator", layout="wide")
st.title("📸 MED PICTURES Word Document Generator")

# Layout options
layout_options = {
    "1 x 2": (1, 2),
    "1 x 3": (1, 3),
    "2 x 2": (2, 2),
    "2 x 3": (2, 3),
    "3 x 2": (3, 2),
    "3 x 3": (3, 3),
}

# Sidebar Inputs
orientation = st.sidebar.selectbox("Select Page Orientation", ["Portrait", "Landscape"])
layout_choice = st.sidebar.selectbox("Select Picture Layout", list(layout_options.keys()))
num_columns, num_rows = layout_options[layout_choice]
pics_per_page = num_columns * num_rows

uploaded_files = st.file_uploader("Upload Pictures", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

project_title = st.text_input("Enter Project Title")
contractor_name = st.text_input("Enter Contractor Name")

preview_placeholder = st.empty()

def generate_preview(files):
    if not files:
        return
    cols = st.columns(num_columns)
    for index, file in enumerate(files):
        col = cols[index % num_columns]
        with col:
            st.image(file, caption=f"Image {index+1}", use_container_width=True)

# Generate Word Document
def create_document():
    doc = Document()

    # Page setup
    section = doc.sections[0]
    if orientation == 'Landscape':
        section.orientation = 1
        section.page_width, section.page_height = section.page_height, section.page_width

    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)

    images = [Image.open(file) for file in uploaded_files]

    # Calculate uniform height
    target_height = 300  # pixels
    resized_images = [img.resize((int(target_height * img.width / img.height), target_height)) for img in images]

    for i in range(0, len(resized_images), pics_per_page):
        if i != 0:
            doc.add_page_break()
        
        # Header
        p = doc.add_paragraph()
        run_red = p.add_run("MED PICTURES: ")
        run_red.font.color.rgb = RGBColor(255, 0, 0)
        run_red.font.size = Pt(14)
        run_black = p.add_run(f"{project_title} by {contractor_name}")
        run_black.font.color.rgb = RGBColor(0, 0, 0)
        run_black.font.size = Pt(14)

        table = doc.add_table(rows=num_rows, cols=num_columns)
        table.autofit = True

        subset = resized_images[i:i+pics_per_page]
        for idx, img in enumerate(subset):
            row = idx // num_columns
            col = idx % num_columns
            cell = table.cell(row, col)

            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format='PNG')
            img_byte_arr.seek(0)
            paragraph = cell.paragraphs[0]
            run = paragraph.add_run()
            run.add_picture(img_byte_arr, width=Inches(2.5))

    doc_stream = io.BytesIO()
    doc.save(doc_stream)
    doc_stream.seek(0)
    return doc_stream

# Show preview
generate_preview(uploaded_files)

if st.button("Generate Word Document"):
    if uploaded_files and project_title and contractor_name:
        doc_stream = create_document()
        st.success("Document generated successfully!")
        st.download_button(
            label="Download Document",
            data=doc_stream,
            file_name=f"MED_PICTURES_{project_title}_by_{contractor_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.warning("Please provide all required inputs.")
