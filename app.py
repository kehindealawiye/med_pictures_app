import streamlit as st
from docx import Document
from docx.shared import Inches, RGBColor
from PIL import Image
import io

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

# Function to generate the Word document
def generate_doc(title, contractor, images, layout, orientation):
    doc = Document()

    # Page orientation and margins
    section = doc.sections[0]
    if orientation == 'Landscape':
        section.orientation = 1  # Landscape
        section.page_width, section.page_height = section.page_height, section.page_width
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)

    rows, cols = layout_options[layout]
    images_per_page = rows * cols

    for i in range(0, len(images), images_per_page):
        # Header
        p = doc.add_paragraph()
        run1 = p.add_run("MED PICTURES: ")
        run1.font.color.rgb = RGBColor(255, 0, 0)
        run1.bold = True
        run2 = p.add_run(f"{title} by {contractor}")
        run2.bold = True

        table = doc.add_table(rows=rows, cols=cols)
        table.autofit = True

        for idx, image_file in enumerate(images[i:i + images_per_page]):
            r, c = divmod(idx, cols)
            cell = table.rows[r].cells[c]
            img = Image.open(image_file)
            img.thumbnail((300, 300))  # Ensure image fits in the cell
            image_stream = io.BytesIO()
            img.save(image_stream, format='PNG')
            image_stream.seek(0)
            cell.paragraphs[0].add_run().add_picture(image_stream, width=Inches(2.2))

        if i + images_per_page < len(images):
            doc.add_page_break()

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# Initialize session state
if "sections" not in st.session_state:
    st.session_state.sections = [0]

# Button to add a new form
new_section = st.button("Generate Another Document")
if new_section:
    st.session_state.sections.append(len(st.session_state.sections))

# Render the forms
for section_id in st.session_state.sections:
    with st.container():
        st.subheader(f"Document Generator #{section_id + 1}")
        with st.form(f"form_{section_id}"):
            col1, col2 = st.columns(2)
            with col1:
                title = st.text_input("Project Title", key=f"title_{section_id}")
                contractor = st.text_input("Contractor Name", key=f"contractor_{section_id}")
                layout = st.selectbox("Select Layout", list(layout_options.keys()), key=f"layout_{section_id}")
            with col2:
                orientation = st.radio("Page Orientation", ["Portrait", "Landscape"], horizontal=True, key=f"orientation_{section_id}")
                images = st.file_uploader("Upload Pictures", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key=f"images_{section_id}")

            submitted = st.form_submit_button("Generate Word Document")

        # Show selected images in grid layout before document generation
        if images:
            st.markdown("### Selected Image Layout Preview")
            rows, cols = layout_options[layout]
            image_count = len(images)
            for i in range(0, image_count, cols):
                cols_preview = st.columns(cols)
                for j in range(cols):
                    if i + j < image_count:
                        img = Image.open(images[i + j])
                        cols_preview[j].image(img, use_container_width=True, caption=f"Image {i+j+1}")
        
        # Generate and download document after form submission
        if submitted:
            if not title or not contractor or not images:
                st.error("Please provide title, contractor name, and upload pictures.")
            else:
                buffer = generate_doc(title, contractor, images, layout, orientation)
                st.success("Document ready!")
                st.download_button(
                    label="📄 Download Word Document",
                    data=buffer,
                    file_name=f"MED_PICTURES_{title.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
