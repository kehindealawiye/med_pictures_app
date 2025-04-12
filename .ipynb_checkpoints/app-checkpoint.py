import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from PIL import Image
import os
from io import BytesIO
from datetime import datetime

def create_document(project_title, contractor_name, images, layout, orientation, image_width, margin_control):
    doc = Document()

    # Set page orientation
    section = doc.sections[0]
    if orientation == 'Landscape':
        section.orientation = 1  # Landscape
        section.page_width, section.page_height = section.page_height, section.page_width

    # Set margins
    if margin_control:
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)

    # Title
    p = doc.add_paragraph()
    run_red = p.add_run("MED PICTURES: ")
    run_red.font.color.rgb = RGBColor(255, 0, 0)
    run_red.font.size = Pt(16)
    run_red.bold = True

    run_black = p.add_run(f"{project_title} by {contractor_name}")
    run_black.font.color.rgb = RGBColor(0, 0, 0)
    run_black.font.size = Pt(16)
    run_black.bold = True

    # Date/time
    doc.add_paragraph(datetime.now().strftime("%d %B %Y, %I:%M %p"))

    # Layout configuration
    columns, rows = map(int, layout.split('x'))
    images_per_page = columns * rows

    # Add images with fixed height alignment
    from docx.shared import Cm
    from PIL import ImageOps

    for idx, img_file in enumerate(images):
        if idx % images_per_page == 0:
            if idx != 0:
                doc.add_page_break()
            # Repeat header
            p = doc.add_paragraph()
            run_red = p.add_run("MED PICTURES: ")
            run_red.font.color.rgb = RGBColor(255, 0, 0)
            run_red.font.size = Pt(16)
            run_red.bold = True

            run_black = p.add_run(f"{project_title} by {contractor_name}")
            run_black.font.color.rgb = RGBColor(0, 0, 0)
            run_black.font.size = Pt(16)
            run_black.bold = True

        image = Image.open(img_file)
        aspect_ratio = image.width / image.height

        # Calculate consistent height based on width
        target_width_px = int(image_width * 96)  # 96 dpi
        target_height_px = int(target_width_px / aspect_ratio)

        image = image.resize((target_width_px, target_height_px))

        buffer = BytesIO()
        image.save(buffer, format='PNG')
        buffer.seek(0)
        doc.add_picture(buffer, width=Inches(image_width))

    # Save to BytesIO
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# Streamlit UI
st.set_page_config(page_title="MED Pictures Generator", layout="wide")
st.title("\U0001F4F8 MED PICTURES Word Document Generator")

project_title = st.text_input("Project Title")
contractor_name = st.text_input("Contractor Name")

images = st.file_uploader("Upload Pictures", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
layout = st.selectbox("Select Layout", ["2x2", "3x2", "3x3"], index=1)
orientation = st.radio("Select Page Orientation", ["Portrait", "Landscape"], index=0)
image_width = st.slider("Image Width (in inches)", 1.0, 6.0, 3.0)
margin_control = st.checkbox("Use 1-inch Margins", value=True)

generate = st.button("Generate Document")

if generate and images:
    doc_stream = create_document(project_title, contractor_name, images, layout, orientation, image_width, margin_control)
    st.download_button(
        label="Download Word Document",
        data=doc_stream,
        file_name=f"MED_PICTURES_{project_title.replace(' ', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
