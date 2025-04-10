import streamlit as st
from docx import Document
from docx.shared import Inches, RGBColor
from io import BytesIO

# Streamlit UI: Collecting user input for layout and image size
st.title("MED PICTURES: Document Generator")

project_title = st.text_input("Project Title")
contractor_name = st.text_input("Contractor Name")

image_files = st.file_uploader("Upload Images", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

# Options for layout and image size
image_size = st.slider("Select Image Width (in inches)", min_value=1, max_value=4, value=3)
layout_option = st.radio("Select Layout", ("2x2", "3x2", "4x2"))

# Button to generate document
if st.button("Generate Document"):
    if project_title and contractor_name and image_files:
        # Create Document
        doc = Document()
        
        # Add Title (left-aligned)
        paragraph = doc.add_paragraph()
        run = paragraph.add_run(f"MED PICTURES: {project_title} by {contractor_name}")
        run.bold = True
        run.font.color.rgb = RGBColor(255, 0, 0)  # Red title
        paragraph.alignment = 0  # Left-aligned
        
        # Determine layout configuration
        if layout_option == "2x2":
            rows, cols = 2, 2
        elif layout_option == "3x2":
            rows, cols = 3, 2
        else:
            rows, cols = 4, 2
        
        # Insert images in selected layout
        image_index = 0
        while image_index < len(image_files):
            table = doc.add_table(rows=rows, cols=cols)
            for row in range(rows):
                for col in range(cols):
                    if image_index < len(image_files):
                        image_stream = BytesIO(image_files[image_index].read())
                        cell = table.cell(row, col)
                        cell.paragraphs[0].add_run().add_picture(image_stream, width=Inches(image_size))
                        image_index += 1
                    else:
                        table.cell(row, col).text = ""  # Empty cell if not enough images
            
            doc.add_paragraph()  # Space after each image table
        
        # Save document and provide download link
        doc_filename = f"MED_PICTURES_{project_title}_{contractor_name}.docx"
        doc_path = f"./{doc_filename}"
        doc.save(doc_path)
        
        st.success(f"Document '{doc_filename}' generated successfully!")
        st.download_button("Download Document", doc_path)
