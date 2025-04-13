import streamlit as st
from PIL import Image
from docx import Document
from docx.shared import Inches, RGBColor
from io import BytesIO
import datetime
import math

st.set_page_config(page_title="MED PICTURES Generator", layout="wide")
st.title("📸 MED PICTURES Word Document Generator")

def generate_doc(images, title, contractor, layout, orientation, margin):
    doc = Document()

    # Orientation
    section = doc.sections[0]
    if orientation == "Landscape":
        section.orientation = 1
        section.page_width, section.page_height = section.page_height, section.page_width

    # Margins
    if margin:
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)

    # Title formatting
    def add_header():
        p = doc.add_paragraph()
        run_red = p.add_run("MED PICTURES: ")
        run_red.font.color.rgb = RGBColor(255, 0, 0)
        run_black = p.add_run(f"{title} by {contractor}")
        run_black.bold = True

    rows, cols = map(int, layout.split("x"))
    per_page = rows * cols

    for i in range(0, len(images), per_page):
        if i > 0:
            doc.add_page_break()
        add_header()

        table = doc.add_table(rows=rows, cols=cols)
        table.autofit = False
        idx = i

        for r in range(rows):
            row = table.rows[r]
            for c in range(cols):
                cell = row.cells[c]
                if idx < len(images):
                    image = Image.open(images[idx])
                    image.thumbnail((300, 300))
                    img_stream = BytesIO()
                    image.save(img_stream, format='PNG')
                    img_stream.seek(0)
                    cell.paragraphs[0].add_run().add_picture(img_stream, width=Inches(2.2))
                    idx += 1

    stream = BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

# Layout Options
layouts = ["1x2", "1x3", "2x2", "2x3", "3x2", "3x3"]

if "generate_count" not in st.session_state:
    st.session_state.generate_count = 1

for gen_id in range(1, st.session_state.generate_count + 1):
    with st.form(f"form_{gen_id}"):
        st.markdown(f"### Document Generator #{gen_id}")
        col1, col2 = st.columns(2)

        with col1:
            title = st.text_input("Project Title", key=f"title_{gen_id}")
            contractor = st.text_input("Contractor Name", key=f"contractor_{gen_id}")
            orientation = st.radio("Page Orientation", ["Portrait", "Landscape"], key=f"orientation_{gen_id}")
            layout = st.selectbox("Layout (rows x columns per page)", layouts, key=f"layout_{gen_id}")
            margin = st.checkbox("Apply 0.5-inch margin", value=True, key=f"margin_{gen_id}")

        with col2:
            uploaded_files = st.file_uploader("Upload Pictures", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key=f"files_{gen_id}")

            # Show layout preview
            if uploaded_files:
                rows, cols = map(int, layout.split("x"))
                st.markdown("**Preview of Selected Images**")
                for r in range(rows):
                    cols_images = uploaded_files[r*cols:(r+1)*cols]
                    cols_stream = st.columns(cols)
                    for i, img in enumerate(cols_images):
                        cols_stream[i].image(img, use_container_width=True)

        submitted = st.form_submit_button("Generate Document")

        if submitted and uploaded_files:
            docx_file = generate_doc(uploaded_files, title, contractor, layout, orientation, margin)
            st.success("Document generated successfully!")
            st.download_button(
                label="📥 Download Word Document",
                data=docx_file,
                file_name=f"MED_PICTURES_{title.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.markdown("---")
            if st.button("Generate Another Document", key=f"another_{gen_id}"):
                st.session_state.generate_count += 1
                st.experimental_rerun()
