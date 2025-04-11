import os
import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import streamlit as st
from PIL import Image

# Function to create a Word document
def create_word_doc(project_title, contractor_name, image_files, image_width, layout, orientation, margin_control):
    doc = Document()

    # Set page orientation
    if orientation == 'Landscape':
        section = doc.sections[0]
        section.orientation = 1  # Landscape
        section.page_width, section.page_height = section.page_height, section.page_width

    # Set margins
    section = doc.sections[0]
    if margin_control:
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)

    # Title section
    title = doc.add_paragraph()
    title_run = title.add_run("MED PICTURES: " + project_title + " by " + contractor_name)
    title_run.font.bold = True
    title_run.font.size = Pt(16)
    title_run.font.color.rgb = RGBColor(255, 0, 0)
    title.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    # Add current date and time
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    doc.add_paragraph("Date: " + current_time, style='Normal')

    # Add images based on the layout choice
    num_images = len(image_files)
    rows = layout[0]
    cols = layout[1]
    image_height = Inches(2.5)  # default height
    image_width = Inches(image_width)

    # Add image placeholders based on selected layout
    table = doc.add_table(rows=rows, cols=cols)
    table.autofit = True

    index = 0
    for row in range(rows):
        for col in range(cols):
            if index < num_images:
                cell = table.cell(row, col)
                img = Image.open(image_files[index])
                img.thumbnail((image_width, image_height))
                img_path = f"temp_image_{index}.png"
                img.save(img_path)
                cell.paragraphs[0].clear()
                cell.paragraphs[0].add_run().add_picture(img_path, width=image_width, height=image_height)
                os.remove(img_path)  # Clean up the temporary image file
                index += 1
            else:
                # Leave the cell empty if there are fewer images
                cell.text = ""

    # Save the document with the correct file name
    save_name = f"MED_PICTURES_{project_title} by {contractor_name}.docx"
    doc.save(save_name)
    return save_name

# Streamlit interface for user input
st.title('Generate MED PICTURES Word Document')

# Project title and contractor name
project_title = st.text_input('Enter Project Title')
contractor_name = st.text_input('Enter Contractor Name')

# Image file input
image_files = st.file_uploader('Upload Images', type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)

# Image width options
image_width = st.slider('Select Image Width', min_value=1, max_value=5, value=3)

# Layout options (rows x columns)
layout = st.selectbox('Select Layout', options=[(2, 2), (3, 2), (3, 3), (4, 4)], index=0)

# Orientation options
orientation = st.selectbox('Select Page Orientation', options=['Portrait', 'Landscape'], index=0)

# Margin control
margin_control = st.checkbox('Enable Custom Margins', value=False)

# Generate button
if st.button('Generate Document'):
    if project_title and contractor_name and image_files:
        # Save images temporarily
        image_paths = []
        for img in image_files:
            with open(f"temp_{img.name}", 'wb') as f:
                f.write(img.getvalue())
            image_paths.append(f"temp_{img.name}")

        # Create the Word document
        try:
            output_path = create_word_doc(project_title, contractor_name, image_paths, image_width, layout, orientation, margin_control)
            st.success(f"Document generated successfully: {output_path}")
            st.download_button('Download Word Document', data=open(output_path, 'rb'), file_name=output_path, mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

        finally:
            # Clean up temporary image files
            for path in image_paths:
                os.remove(path)
    else:
        st.error('Please provide all inputs.')
