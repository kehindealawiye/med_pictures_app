import streamlit as st
import docx
from docx.shared import Inches, RGBColor
from io import BytesIO
from datetime import datetime
import os
from PIL import Image
import math

# Page title and layout settings
st.set_page_config(page_title="MED Pictures Generator", layout="wide")
st.title("📸 MED PICTURES Word Document Generator")

# Collect user inputs
project_title = st.text_input("Project Title")
contractor_name = st.text_input("Contractor Name")
orientation = st.selectbox("Select Page Orientation", ["Portrait", "Landscape"])
num_pictures = st.number_input("Number of Pictures", min_value=1, max_value=12, step=1)
layout = st.selectbox("Select Layout", ["1x1", "2x2", "3x2", "3x3"])
image_width = st.selectbox("Select Image Width (inches)", [1, 2, 3, 4, 5])
margin_control = st.checkbox("Enable Margin Control")

# Upload pictures
uploaded_pics = []
for i in range(num_pictures):
    uploaded_pics.append(st.file_uploader(f"Upload Picture {i+1}", type=["jpg", "png", "jpeg"], key=f"pic_{i+1}"))

# Define the document generation function
def create_document():
    doc = docx.Document()
    section = doc.sections[0]

    # Set page orientation
    if orientation == 'Landscape':
        section.orientation = 1  # Landscape
        section.page_width, section.page_height = section.page_height, section.page_width

    # Set margins if margin_control is checked
    if margin_control:
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)

    # Title with different color for "MED PICTURES"
    title = doc.add_paragraph()
    title.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT
    run = title.add_run(f"MED PICTURES: {project_title} by {contractor_name}")
    run.font.size = docx.shared.Pt(14)
    run.font.color.rgb = RGBColor(255, 0, 0)  # Red color for "MED PICTURES"
    title.add_run(f" {contractor_name}")
    title.paragraph_format.space_after = Inches(0.3)

    # Add the date and time
    date_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    doc.add_paragraph(f"Generated on: {date_time}")

    # Add pictures
    pictures = []
    for file in uploaded_pics:
        if file:
            picture = Image.open(file)
            pictures.append(picture)

    # Image layout and resizing
    layout_map = {
        "1x1": (1, 1),
        "2x2": (2, 2),
        "3x2": (3, 2),
        "3x3": (3, 3),
    }

    rows, cols = layout_map.get(layout, (1, 1))
    images_per_row = cols
    image_height = Inches(2)  # Fixed image height

    # Create the picture grid
    for i in range(0, len(pictures), images_per_row):
        row_pictures = pictures[i:i+images_per_row]
        row = doc.add_paragraph()
        for pic in row_pictures:
            pic_path = BytesIO()
            pic = pic.resize((int(image_width * 100), int(image_height * 100)))  # Resize image
            pic.save(pic_path, format="PNG")
            pic_path.seek(0)

            run = row.add_run()
            run.add_picture(pic_path, width=Inches(image_width), height=image_height)
            run.add_run("\t")  # Add some space between images
        row.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT

    # Ensure the header appears on each page
    header = doc.sections[0].header
    paragraph = header.paragraphs[0]
    paragraph.text = f"MED PICTURES: {project_title} by {contractor_name}"

    # Save to BytesIO instead of disk
    doc_stream = BytesIO()
    doc.save(doc_stream)
    doc_stream.seek(0)
    
    return doc_stream

# Button to generate the document
if st.button("Generate Word Document"):
    doc_stream = create_document()
    st.download_button(
        label="Download Document",
        data=doc_stream,
        file_name="MED_PICTURES_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
