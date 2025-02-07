import streamlit as st
import docx
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image
import os
from io import BytesIO

st.title("📄 PDF Merger Tool")
st.write("Upload multiple files, reorder them, and merge into a single PDF.")

# Upload files
uploaded_files = st.file_uploader(
    "Upload Files", type=["txt", "docx", "pdf", "jpg", "png", "xlsx"], accept_multiple_files=True
)

if uploaded_files:
    file_data = [{"Filename": f.name} for f in uploaded_files]

    # ✅ FIX: Use `st.data_editor` for proper drag-and-drop reordering (instead of checkboxes)
    file_df = st.data_editor(
        file_data, num_rows="dynamic", use_container_width=True
    )

    # Get the reordered filenames
    reordered_files = [row["Filename"] for row in file_df]

    if st.button("Merge Files"):
        pdf_docs = []
        merged_text = ""

        for filename in reordered_files:
            file = next((f for f in uploaded_files if f.name == filename), None)
            if file:
                ext = os.path.splitext(file.name)[1].lower()

                if ext == ".txt":
                    merged_text += file.read().decode("utf-8") + "\n\n"

                elif ext == ".docx":
                    doc = docx.Document(file)
                    doc_text = "\n".join([para.text for para in doc.paragraphs])

                    # Convert `.docx` text to PDF
                    text_pdf = fitz.open()
                    text_page = text_pdf.new_page(width=595, height=842)
                    text_rect = fitz.Rect(50, 50, 545, 800)
                    text_page.insert_textbox(text_rect, doc_text, fontsize=12, fontname="helv")
                    pdf_docs.append(text_pdf)

                elif ext == ".xlsx":
                    df = pd.read_excel(file)
                    excel_text = df.to_string()

                    # Convert `.xlsx` text to PDF
                    text_pdf = fitz.open()
                    text_page = text_pdf.new_page(width=595, height=842)
                    text_rect = fitz.Rect(50, 50, 545, 800)
                    text_page.insert_textbox(text_rect, excel_text, fontsize=12, fontname="helv")
                    pdf_docs.append(text_pdf)

                elif ext == ".pdf":
                    pdf_docs.append(fitz.open(stream=file.read(), filetype="pdf"))

                elif ext in (".jpg", ".png"):
                    img = Image.open(file)
                    img_bytes = BytesIO()
                    img.save(img_bytes, format="PDF")
        
