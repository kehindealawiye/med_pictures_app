import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.section import WD_ORIENT
from PIL import Image
from io import BytesIO
import datetime

st.set_page_config(page_title="MED Pictures Generator", layout="wide")
st.title("\U0001F4F8 MED PICTURES Word Document Generator")

# Layout options
layout_options = {
    "1 x 2": (1, 2),
    "1 x 3": (1, 3),
    "2 x 2": (2, 2),
    "2 x 3": (2, 3),
    "3 x 2": (3, 2),
    "3 x 3": (3, 3)
}

if "forms" not in st.session_state:
    st.session_state.forms = [0]

# Handle form submissions and document generations
def handle_form(idx):
    with st.form(key=f"form_{idx}"):
        st.subheader(f"Document Options #{idx + 1}")

        project_title = st.text_input("Project Title", key=f"project_title_{idx}")
        contractor_name = st.text_input("Contractor Name", key=f"contractor_name_{idx}")
        layout = st.selectbox("Select Layout", list(layout_options.keys()), key=f"layout_{idx}")
        orientation = st.radio("Page Orientation", ["Portrait", "Landscape"], key=f"orientation_{idx}")
        images = st.file_uploader("Upload Images", accept_multiple_files=True, type=['jpg', 'jpeg', 'png'], key=f"images_{idx}")

        submitted = st.form_submit_button("Generate Document")

        if submitted and project_title and contractor_name and images:
            cols, rows = layout_options[layout]
            doc = Document()

            # Orientation
            section = doc.sections[0]
            if orientation == 'Landscape':
                section.orientation = WD_ORIENT.LANDSCAPE
                section.page_width, section.page_height = section.page_height, section.page_width

            # Margins
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)

            # Header style
            def add_header():
                paragraph = doc.add_paragraph()
                run_red = paragraph.add_run("MED PICTURES: ")
                run_red.bold = True
                run_red.font.size = Pt(16)
                run_red.font.color.rgb = RGBColor(255, 0, 0)
                run_black = paragraph.add_run(f"{project_title} by {contractor_name}")
                run_black.font.size = Pt(16)
                run_black.font.color.rgb = RGBColor(0, 0, 0)

            images_per_page = cols * rows
            for i in range(0, len(images), images_per_page):
                if i != 0:
                    doc.add_page_break()
                add_header()

                table = doc.add_table(rows=rows, cols=cols)
                table.autofit = True

                chunk = images[i:i + images_per_page]
                for idx_img, img_file in enumerate(chunk):
                    row_idx = idx_img // cols
                    col_idx = idx_img % cols
                    cell = table.cell(row_idx, col_idx)
                    img = Image.open(img_file)
                    img.thumbnail((400, 400))
                    buffered = BytesIO()
                    img.save(buffered, format="PNG")
                    cell.paragraphs[0].add_run().add_picture(buffered, width=Inches(2.2))

            buffer = BytesIO()
            doc.save(buffer)
            buffer.seek(0)

            st.download_button(
                label="Download Word Document",
                data=buffer,
                file_name=f"MED_PICTURES_{project_title}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

            st.success("Document generated successfully!")

            if st.button("Generate Another Document", key=f"another_{idx}"):
                st.session_state.forms.append(len(st.session_state.forms))

        # --- Preview Section ---
        if images:
            st.subheader("\U0001F5BC Selected Image Layout Preview")
            cols_preview = layout_options[layout][0]
            rows_preview = layout_options[layout][1]
            for row in range((len(images) + cols_preview - 1) // cols_preview):
                cols = st.columns(cols_preview)
                for col_index in range(cols_preview):
                    image_index = row * cols_preview + col_index
                    if image_index < len(images):
                        img_file = images[image_index]
                        img = Image.open(img_file)
                        with cols[col_index]:
                            st.image(img, caption=f"Image {image_index + 1}", use_container_width=True)

# Render all forms
for idx in st.session_state.forms:
    handle_form(idx)