import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from PIL import Image
import io
from datetime import datetime

# Page config
st.set_page_config(page_title="MED Pictures Generator", layout="wide")
st.title("📸 MED PICTURES Word Document Generator")

# Layout options
layouts = {
    "1 x 2": (1, 2),
    "1 x 3": (1, 3),
    "2 x 2": (2, 2),
    "2 x 3": (2, 3),
    "3 x 2": (3, 2),
    "3 x 3": (3, 3),
}

def render_form(key_prefix=""):
    with st.form(key=f"form_{key_prefix}"):
        project_title = st.text_input("Project Title", key=f"title_{key_prefix}")
        contractor_name = st.text_input("Contractor Name", key=f"contractor_{key_prefix}")
        orientation = st.selectbox("Page Orientation", ["Portrait", "Landscape"], key=f"orient_{key_prefix}")
        layout = st.selectbox("Select Layout (Columns x Rows per Page)", list(layouts.keys()), key=f"layout_{key_prefix}")
        images = st.file_uploader("Upload Pictures", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key=f"upload_{key_prefix}")

        if images:
            st.markdown("**Selected Images Preview**")
            cols, rows = layouts[layout]
            grid_cols = st.columns(cols)
            for i, img_file in enumerate(images):
                with grid_cols[i % cols]:
                    st.image(img_file, use_container_width=True)

        submitted = st.form_submit_button("Generate Document")
        if submitted and project_title and contractor_name and images:
            docx_file = create_doc(project_title, contractor_name, orientation, layout, images)
            st.success("Document generated successfully!")
            st.download_button("📄 Download Word Document", data=docx_file.getvalue(), file_name="MED_Pictures.docx")
            st.markdown("---")
            st.button("Generate Another Document", key=f"gen_another_{key_prefix}", on_click=lambda: st.session_state.update({f"show_{key_prefix}next": True}))

def create_doc(title, contractor, orientation, layout, images):
    cols, rows = layouts[layout]
    images_per_page = cols * rows
    doc = Document()

    # Page orientation and margin
    section = doc.sections[0]
    if orientation == "Landscape":
        section.orientation = 1
        section.page_width, section.page_height = section.page_height, section.page_width
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)

    # Resize params
    max_width = section.page_width - section.left_margin - section.right_margin
    max_height = section.page_height - section.top_margin - section.bottom_margin - Inches(0.5)
    image_width = max_width / cols - Inches(0.1)
    image_height = max_height / rows - Inches(0.1)

    for idx, image_file in enumerate(images):
        if idx % images_per_page == 0:
            if idx != 0:
                doc.add_page_break()
            add_header(doc, title, contractor)

        img = Image.open(image_file)
        img.thumbnail((int(image_width * 100), int(image_height * 100)))
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        buf.seek(0)

        if idx % cols == 0:
            table = doc.add_table(rows=1, cols=cols)
            table.autofit = False
            row_cells = table.rows[0].cells

        cell_idx = idx % cols
        cell = row_cells[cell_idx]
        run = cell.paragraphs[0].add_run()
        run.add_picture(buf, width=image_width)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

def add_header(doc, title, contractor):
    para = doc.add_paragraph()
    run_red = para.add_run("MED PICTURES: ")
    run_red.font.color.rgb = RGBColor(255, 0, 0)
    run_red.bold = True
    run_red.font.size = Pt(14)

    run_black = para.add_run(f"{title} by {contractor}")
    run_black.font.color.rgb = RGBColor(0, 0, 0)
    run_black.font.size = Pt(14)

# Render first form
render_form("1")

# If user clicks Generate Another Document, show another form
if st.session_state.get("show_1next"):
    render_form("2")
if st.session_state.get("show_2next"):
    render_form("3")