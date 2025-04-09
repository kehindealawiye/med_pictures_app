import streamlit as st
from docx import Document
from docx.shared import Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image
import io

st.title("Project Monitoring Report Generator")

# Inputs
project_title = st.text_input("Project Title")
contractor_name = st.text_input("Contractor Name")
uploaded_files = st.file_uploader("Upload Project Images", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

if st.button("Generate Word Document") and uploaded_files:
    doc = Document()

    # Red Title
    red_title = doc.add_paragraph()
    run = red_title.add_run("MED PICTURES:")
    run.font.color.rgb = RGBColor(255, 0, 0)
    red_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Black Subtitle
    subtitle = doc.add_paragraph(f"{project_title} by {contractor_name}")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Image Grid in 2 Columns
    table = doc.add_table(rows=0, cols=2)
    row = None

    for i, file in enumerate(uploaded_files):
        if i % 2 == 0:
            row = table.add_row().cells

        image = Image.open(file)
        caption = file.name.split('.')[0].replace('_', ' ')
        img_stream = io.BytesIO()
        image.save(img_stream, format='PNG')
        img_stream.seek(0)

        cell = row[i % 2]
        paragraph = cell.paragraphs[0]
        run = paragraph.add_run()
        run.add_picture(img_stream, width=Inches(2.5))
        paragraph.add_run(f"\n{caption}")

    # Save and Export
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    filename = f'MED PICTURES: "{project_title}" by "{contractor_name}".docx'

    st.success("Word document generated!")
    st.download_button(
        label="Download Word Report",
        data=buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
