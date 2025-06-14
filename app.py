from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import os
import streamlit as st
from PIL import Image
from formatter import insert_images_ai_style

# Add custom CSS for RTL support
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
        .stMarkdown {
            direction: rtl;
            text-align: right;
        }
    </style>
""", unsafe_allow_html=True)

# ממשק משתמש
st.title("📄 Word Image Formatter (AI Mode)")

uploaded_file = st.file_uploader("העלה קובץ Word (.doc או .docx)", type=["doc", "docx"])
uploaded_images = st.file_uploader("העלה תמונות", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

if st.button("שבץ את התמונות לסוף הדוח"):
    if not uploaded_file or not uploaded_images:
        st.error("יש להעלות גם קובץ Word וגם תמונות")
    else:
        try:
            # יצירת תיקיות
            os.makedirs("input/images", exist_ok=True)
            os.makedirs("output", exist_ok=True)

            # שמירת קובץ Word
            input_path = os.path.join("input", uploaded_file.name)
            with open(input_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # המרה ל-DOCX אם צריך
            if input_path.endswith(".doc"):
                input_path = convert_doc_to_docx(input_path)

            # שמירת תמונות
            for img in uploaded_images:
                img_path = os.path.join("input/images", img.name)
                image = Image.open(img)
                image.save(img_path, optimize=True, quality=85)

            # עיבוד הדו"ח
            output_path = os.path.join("output", "ready_report.docx")
            final_path = insert_images_ai_style(input_path, "input/images", output_path)

            # קריאת הקובץ להורדה
            with open(final_path, "rb") as f:
                report_data = f.read()

            st.success("📄 הדוח הופק בהצלחה!")
            st.download_button("הורד", data=report_data, file_name="ready_report.docx")

        except Exception as e:
            st.error(f"אירעה שגיאה: {str(e)}")

        # ניקוי רק אחרי עיבוד והורדה
        finally:
            import shutil
            if os.path.exists("input"):
                shutil.rmtree("input")
            if os.path.exists("output"):
                shutil.rmtree("output")
