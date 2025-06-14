import os
import shutil
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import streamlit as st

st.set_page_config(page_title="Word Report Generator", layout="centered")

# Insert images into Word doc
def insert_images_ai_style(doc_path, images_folder, output_path):
    doc = Document(doc_path)
    section = doc.sections[0]
    section.right_margin = Inches(1)
    section.left_margin = Inches(1)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)

    for paragraph in doc.paragraphs:
        paragraph.paragraph_format.right_to_left = True

    images = sorted(os.listdir(images_folder))
    img_idx = 0
    total_images = len(images)
    page_number = 1

    while img_idx < total_images:
        table = doc.add_table(rows=2, cols=2)
        table.autofit = True
        image_number = 1

        for row in range(2):
            for col in range(2):
                if img_idx >= total_images:
                    break
                img_path = os.path.join(images_folder, images[img_idx])
                cell = table.cell(row, col)
                paragraph = cell.paragraphs[0]
                paragraph.paragraph_format.right_to_left = True
                run = paragraph.add_run()
                run.add_picture(img_path, width=Inches(2.5))
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

                caption = cell.add_paragraph(f"{page_number}.{image_number}")
                caption.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                caption.runs[0].font.size = Pt(10)
                caption.paragraph_format.right_to_left = True

                img_idx += 1
                image_number += 1

        doc.add_paragraph("\n")
        desc_table = doc.add_table(rows=1, cols=1)
        desc_cell = desc_table.cell(0, 0)
        desc_cell.height = Inches(1.0)

        tcBorders = OxmlElement("w:tcBorders")
        for side in ['top', 'bottom', 'left', 'right']:
            border = OxmlElement(f"w:{side}")
            border.set(qn("w:val"), "single")
            border.set(qn("w:sz"), "4")
            border.set(qn("w:space"), "0")
            border.set(qn("w:color"), "000000")
            tcBorders.append(border)
        desc_cell._tc.get_or_add_tcPr().append(tcBorders)

        set_cell_rtl(desc_cell)
        desc_paragraph = desc_cell.paragraphs[0]
        desc_paragraph.paragraph_format.right_to_left = True
        desc_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        desc_paragraph.add_run("תיאור תמונות:")

        doc.add_paragraph("")
        page_number += 1

        if img_idx < total_images:
            doc.add_page_break()

    doc.save(output_path)
    return output_path

# Apply RTL for a cell
def set_cell_rtl(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    tcPr.append(bidi)

# CSS for RTL in text inputs only (keep UI in English)
st.markdown("""
    <style>
        .stTextInput > div > div > input {
            direction: rtl;
            text-align: right;
        }
        .stTextArea > div > div > textarea {
            direction: rtl;
            text-align: right;
        }
    </style>
""", unsafe_allow_html=True)

st.markdown("## 📄 Word Report Generator", unsafe_allow_html=True)

if 'restart' not in st.session_state:
    st.session_state.restart = False

uploaded_file = st.file_uploader("Upload Word file (.docx only)", type=["docx"])
uploaded_images = st.file_uploader("Upload images", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

if uploaded_images:
    # Arrange images into 3 rows and dynamic number of columns, filled bottom to top, right to left
    num_rows = 3
    num_images = len(uploaded_images)
    num_cols = (num_images + num_rows - 1) // num_rows

    # Fill columns bottom to top
    grid = [[] for _ in range(num_cols)]
    for idx, img in enumerate(uploaded_images):
        col_idx = idx // num_rows
        grid[col_idx].insert(0, img)  # insert at top

    cols = st.columns(num_cols)[::-1]  # reverse the columns order for right-to-left filling
    for col, images in zip(cols, grid[::-1]):
        with col:
            for img in images:
                st.image(img, width=160)
                st.caption(img.name)

if uploaded_file and uploaded_file.name.endswith(".doc"):
    st.error("⚠️ .doc files are not supported. Please convert to .docx and try again.")
    st.stop()

if st.button("Generate Report"):
    if not uploaded_file or not uploaded_images:
        st.error("Please upload both a Word file and images.")
    else:
        try:
            if os.path.exists("input"):
                shutil.rmtree("input")
            if os.path.exists("output"):
                shutil.rmtree("output")
            os.makedirs("input/images", exist_ok=True)
            os.makedirs("output", exist_ok=True)

            input_path = os.path.join("input", uploaded_file.name)
            with open(input_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            for idx, img in enumerate(uploaded_images):
                safe_name = img.name.encode('utf-8', 'ignore').decode('utf-8', 'ignore')
                img_path = os.path.join("input/images", safe_name)
                with open(img_path, "wb") as f:
                    f.write(img.getbuffer())

            output_path = os.path.join("output", "ready_report.docx")
            final_path = insert_images_ai_style(input_path, "input/images", output_path)

            with open(final_path, "rb") as f:
                report_data = f.read()

            st.success("✅ Report generated successfully!")
            st.download_button("Download Report", data=report_data, file_name="ready_report.docx")

            if st.button("Generate another report"):
                st.session_state.restart = True

        except Exception as e:
            st.error(f"Error: {str(e)}")

if st.session_state.restart:
    st.session_state.restart = False
    st.experimental_rerun()
