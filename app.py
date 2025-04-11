import streamlit as st
from docx import Document
from docx.shared import Inches, RGBColor
from io import BytesIO

st.title("MED PICTURES: Document Generator")

project_title = st.text_input("Project Title")
contractor_name = st.text_input("Contractor Name")

image_files = st.file_uploader("Upload Images", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

image_size = st.slider("Select Image Width (in inches)", min_value=1, max_value=4, value=3)
layout_option = st.radio("Select Layout", ("2x2", "3x2", "4x2"))

if st.button("Generate Document"):
    if project_title and contractor_name and image_files:
        doc = Document()
        
        # Title: red, bold, left-aligned
        paragraph = doc.add_paragraph()
        run = paragraph.add_run(f"MED PICTURES: {project_title} by {contractor_name}")
        run.bold = True
        run.font.color.rgb = RGBColor(255, 0, 0)
        paragraph.alignment = 0
        
        # Layout settings
        if layout_option == "2x2":
            rows, cols = 2, 2
        elif layout_option == "3x2":
            rows, cols = 3, 2
        else:
            rows, cols = 4, 2

        image_index = 0
        total_images = len(image_files)

        while image_index < total_images:
            table = doc.add_table(rows=rows, cols=cols)
            for r in range(rows):
                for c in range(cols):
                    if image_index < total_images:
                        file = image_files[image_index]
                        file.seek(0)  # Rewind the file
                        image_stream = BytesIO(file.read())
                        cell = table.cell(r, c)
                        cell.paragraphs[0].add_run().add_picture(image_stream, width=Inches(image_size))
                        image_index += 1
                    else:
                        table.cell(r, c).text = ""
            doc.add_paragraph()  # space after table

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        filename = f"MED_PICTURES_{project_title} by {contractor_name}.docx"
        st.success(f"Document created: {filename}")
        st.download_button("Download Word Document", data=buffer, file_name=filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    else:
        st.warning("Please fill in all fields and upload at least one image.")
