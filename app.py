import streamlit as st
from PIL import Image
from io import BytesIO
from docx import Document
from docx.shared import Inches
import datetime

st.set_page_config(page_title="MED PICTURES Generator", layout="wide")
st.title("📸 MED PICTURES Word Document Generator")

def get_image_from_upload(uploaded_file):
    image = Image.open(uploaded_file)
    return image

def insert_images_in_grid(doc, images, rows, cols, orientation):
    from docx.shared import Cm
    section = doc.sections[-1]
    table = doc.add_table(rows=rows, cols=cols)
    table.autofit = True

    pic_index = 0
    total_images = len(images)

    for r in range(rows):
        row = table.rows[r]
        for c in range(cols):
            if pic_index >= total_images:
                break
            cell = row.cells[c]
            paragraph = cell.paragraphs[0]
            run = paragraph.add_run()
            image_stream = BytesIO()
            images[pic_index].save(image_stream, format='PNG')
            image_stream.seek(0)
            run.add_picture(image_stream, width=Cm(6))
            pic_index += 1
    return doc

def create_document(title, contractor, layout, orientation, images):
    doc = Document()

    # Set orientation and margins
    section = doc.sections[0]
    if orientation == "Landscape":
        section.orientation = 1  # Landscape
        section.page_width, section.page_height = section.page_height, section.page_width
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)

    rows, cols = layout
    max_per_page = rows * cols
    total_pages = -(-len(images) // max_per_page)  # Ceiling division

    for page_num in range(total_pages):
        if page_num > 0:
            doc.add_page_break()
        doc.add_paragraph().add_run().add_break()

        header_para = doc.add_paragraph()
        run_red = header_para.add_run("MED PICTURES: ")
        run_red.bold = True
        run_red.font.color.rgb = docx.shared.RGBColor(255, 0, 0)
        header_para.add_run(f"{title} by {contractor}")

        start = page_num * max_per_page
        end = min(start + max_per_page, len(images))
        page_images = images[start:end]
        insert_images_in_grid(doc, page_images, rows, cols, orientation)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

layout_options = {
    "1 x 2": (2, 1),
    "1 x 3": (3, 1),
    "2 x 2": (2, 2),
    "2 x 3": (3, 2),
    "3 x 2": (2, 3),
    "3 x 3": (3, 3)
}

if "doc_counter" not in st.session_state:
    st.session_state.doc_counter = 0

with st.form(key=f"form_{st.session_state.doc_counter}"):
    title = st.text_input("Project Title")
    contractor = st.text_input("Contractor Name")
    layout_key = st.selectbox("Select layout", list(layout_options.keys()))
    orientation = st.radio("Select Page Orientation", ["Portrait", "Landscape"])
    uploaded_files = st.file_uploader("Upload Pictures", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

    preview_images = [get_image_from_upload(f) for f in uploaded_files] if uploaded_files else []
    layout = layout_options[layout_key]
    rows, cols = layout

    if preview_images:
        st.markdown("### Preview:")
        for i in range(0, len(preview_images), cols):
            cols_layout = st.columns(cols)
            for j in range(cols):
                if i + j < len(preview_images):
                    with cols_layout[j]:
                        st.image(preview_images[i + j], use_container_width=True)

    submitted = st.form_submit_button("Generate Document")

    if submitted and uploaded_files:
        doc_file = create_document(title, contractor, layout, orientation, preview_images)
        st.download_button(
            label="📥 Download Word Document",
            data=doc_file,
            file_name=f"MED_PICTURES_{title}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        st.success("Document generated successfully!")

        if st.button("Generate Another Document"):
            st.session_state.doc_counter += 1