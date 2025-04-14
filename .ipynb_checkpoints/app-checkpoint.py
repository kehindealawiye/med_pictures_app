# === IMPORT LIBRARIES ===
import streamlit as st
from docx import Document
from docx.shared import Inches, RGBColor
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from PIL import Image
import io

# ✅ Set page config before any Streamlit output
st.set_page_config(page_title="MED Pictures Generator", layout="wide")

# === TITLE ===
st.title("📸 MED PICTURES Word Document Generator")

# === LAYOUT OPTIONS MAPPING ===
layout_options = {
    "1 x 2": (1, 2),
    "1 x 3": (1, 3),
    "2 x 2": (2, 2),
    "2 x 3": (2, 3),
    "3 x 2": (3, 2),
    "3 x 3": (3, 3),
}

# === DOCUMENT GENERATION FUNCTION ===
def generate_doc(title, contractor, images, layout, orientation):
    rows, cols = layout_options[layout]
    images_per_page = rows * cols

    # Create Word document
    doc = Document()
    section = doc.sections[0]

    # Set orientation
    if orientation == 'Landscape':
        section.orientation = 1
        section.page_width, section.page_height = section.page_height, section.page_width

    # Set minimal page margins
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)

    # Process images and paginate
    for i in range(0, len(images), images_per_page):
        # Add project header once per page
        p = doc.add_paragraph()
        run1 = p.add_run("MED PICTURES: ")
        run1.font.color.rgb = RGBColor(255, 0, 0)
        run1.bold = True
        run2 = p.add_run(f"{title} by {contractor}")
        run2.bold = True

        # Add table for image layout
        table = doc.add_table(rows=rows, cols=cols)
        table.autofit = False

        # Center-align all cells
        for row in table.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # Insert images
        for idx, image_file in enumerate(images[i:i + images_per_page]):
            r, c = divmod(idx, cols)
            cell = table.rows[r].cells[c]

            img = Image.open(image_file)
            img.thumbnail((600, 600))  # Resize for uniform height
            image_stream = io.BytesIO()
            img.save(image_stream, format='PNG')
            image_stream.seek(0)

            paragraph = cell.paragraphs[0]
            run = paragraph.add_run()
            run.add_picture(image_stream, width=Inches(2.2))  # Adjust width

        # Page break if more images
        if i + images_per_page < len(images):
            doc.add_page_break()

    # Return Word file in memory
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# === USER INPUT FORM ===
with st.form("image_form"):
    st.subheader("🔧 Document Configuration")

    title = st.text_input("Project Title")
    contractor = st.text_input("Contractor Name")
    orientation = st.selectbox("Page Orientation", ["Portrait", "Landscape"])
    layout = st.selectbox("Image Layout", list(layout_options.keys()))
    uploaded_images = st.file_uploader("Upload Images", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

    submitted = st.form_submit_button("Generate Word Document")

# === DOCUMENT GENERATION AND DOWNLOAD ===
if submitted:
    if not title or not contractor or not uploaded_images:
        st.error("Please provide all required inputs.")
    else:
        word_file = generate_doc(title, contractor, uploaded_images, layout, orientation)
        st.success("✅ Document ready!")
        st.download_button(
            label="📥 Download Document",
            data=word_file,
            file_name="MED_Pictures.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
