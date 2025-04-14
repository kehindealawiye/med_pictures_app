import streamlit as st
from docx import Document
from docx.shared import Inches, RGBColor
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from PIL import Image
import io

# --- Set page config ---
st.set_page_config(page_title="📸 MED Pictures Generator", layout="wide")
st.title(u"\U0001F4F8 MED PICTURES Word Document Generator")  # 📸

# --- Crop size mapping (in inches) ---
crop_options = {
    "Portrait Short (11.03 x 9 cm)": (3.44, 2.83),
    "Portrait Tall (13.26 x 9.3 cm)": (4.15, 2.95),
    "Landscape (8.56 x 10.58 cm)": (2.83, 4.17),
    "Wide Landscape (8.56 x 18.94 cm)": (2.83, 7.46),
}

# --- Helper: pad image to size with white background ---
def pad_image_to_size(img, target_size, color=(255, 255, 255)):
    img.thumbnail(target_size, Image.Resampling.LANCZOS)
    new_img = Image.new("RGB", target_size, color)
    left = (target_size[0] - img.width) // 2
    top = (target_size[1] - img.height) // 2
    new_img.paste(img, (left, top))
    return new_img

# --- User input form ---
with st.form("image_form"):
    st.subheader("🛠️ Document Configuration")
    title = st.text_input("Project Title")
    contractor = st.text_input("Contractor Name")
    orientation = st.selectbox("Page Orientation", ["Portrait", "Landscape"])

    uploaded_images = st.file_uploader("Upload Images", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

    crop_selections = []
    if uploaded_images:
        st.markdown("### ✂️ Select Crop Size per Image")
        for idx, image_file in enumerate(uploaded_images):
            col1, col2 = st.columns([1, 2])
            with col1:
                st.image(image_file, caption=f"Image {idx+1}", use_container_width=True)
            with col2:
                crop = st.selectbox(
                    f"Crop size for Image {idx+1}",
                    options=list(crop_options.keys()),
                    key=f"crop_{idx}"
                )
                crop_selections.append(crop)

    layout = st.selectbox("Grid Layout per Page", ["1 x 2", "1 x 3", "2 x 2", "2 x 3", "3 x 2", "3 x 3"])

    submitted = st.form_submit_button("📄 Generate Word Document")

# --- Document generator ---
def generate_doc(title, contractor, images, crop_sizes, layout, orientation):
    layout_map = { "1 x 2": (1, 2), "1 x 3": (1, 3), "2 x 2": (2, 2), "2 x 3": (2, 3), "3 x 2": (3, 2), "3 x 3": (3, 3) }
    rows, cols = layout_map[layout]
    images_per_page = rows * cols

    # Group by crop type
    grouped = {}
    for img, crop in zip(images, crop_sizes):
        grouped.setdefault(crop, []).append(img)
    
    ordered_images = []
    ordered_crop_sizes = []
    for crop, imgs in grouped.items():
        ordered_images.extend(imgs)
        ordered_crop_sizes.extend([crop] * len(imgs))

    doc = Document()
    section = doc.sections[0]

    if orientation == 'Landscape':
        section.orientation = 1
        section.page_width, section.page_height = section.page_height, section.page_width

    section.top_margin = Inches(0.3)
    section.bottom_margin = Inches(0.3)
    section.left_margin = Inches(0.4)
    section.right_margin = Inches(0.4)

    usable_width = section.page_width - section.left_margin - section.right_margin
    col_width = usable_width / cols

    for i in range(0, len(ordered_images), images_per_page):
        # Header
        p = doc.add_paragraph()
        run1 = p.add_run("MED PICTURES: ")
        run1.font.color.rgb = RGBColor(255, 0, 0)
        run1.bold = True
        run2 = p.add_run(f"{title} by {contractor}")
        run2.bold = True

        # Layout table
        table = doc.add_table(rows=rows, cols=cols)
        table.autofit = False

        for row in table.rows:
            for cell in row.cells:
                cell.width = col_width
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        for idx, image_file in enumerate(ordered_images[i:i + images_per_page]):
            r, c = divmod(idx, cols)
            cell = table.rows[r].cells[c]

            img = Image.open(image_file).convert("RGB")
            crop_label = ordered_crop_sizes[i + idx]
            width_in, height_in = crop_options[crop_label]

            target_px = (int(width_in * 96), int(height_in * 96))
            img = pad_image_to_size(img, target_px)

            img_stream = io.BytesIO()
            img.save(img_stream, format='PNG')
            img_stream.seek(0)

            paragraph = cell.paragraphs[0]
            run = paragraph.add_run()
            run.add_picture(img_stream, width=Inches(width_in))

        if i + images_per_page < len(ordered_images):
            doc.add_page_break()

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- Document output ---
if submitted:
    if not title or not contractor or not uploaded_images:
        st.error("Please provide all inputs.")
    else:
        word_doc = generate_doc(title, contractor, uploaded_images, crop_selections, layout, orientation)
        st.success("✅ Document ready!")
        st.download_button("📥 Download Word Document", word_doc, "MED_PICTURES.docx")
