import streamlit as st
import docx
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image
import os
from io import BytesIO

st.title("📄 PDF Merger Tool")
st.write("Upload multiple files, reorder them, and merge into a single PDF.")

uploaded_files = st.file_uploader("Upload Files", type=["txt", "docx", "pdf", "jpg", "png", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    file_order = [f.name for f in uploaded_files]
    reordered_files = st.multiselect("Drag to reorder files", file_order, default=file_order)

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
                    merged_text += "\n".join([para.text for para in doc.paragraphs]) + "\n\n"
                elif ext == ".pdf":
                    pdf_docs.append(fitz.open(stream=file.read(), filetype="pdf"))
                elif ext in (".jpg", ".png"):
                    img = Image.open(file)
                    img_bytes = BytesIO()
                    img.save(img_bytes, format="PDF")
                    pdf_docs.append(fitz.open("pdf", img_bytes.getvalue()))
                elif ext == ".xlsx":
                    df = pd.read_excel(file)
                    merged_text += df.to_string() + "\n\n"

        if merged_text:
            text_pdf = fitz.open()
            text_page = text_pdf.new_page(width=595, height=842)
            text_rect = fitz.Rect(50, 50, 545, 800)
            text_page.insert_textbox(text_rect, merged_text, fontsize=12, fontname="helv")
            pdf_docs.insert(0, text_pdf)

        merged_doc = fitz.open()
        for pdf in pdf_docs:
            merged_doc.insert_pdf(pdf)

        output_path = "merged_output.pdf"
        merged_doc.save(output_path)

        st.success("✅ Merged PDF is ready!")
        st.download_button("Download Merged PDF", open(output_path, "rb"), file_name="merged_output.pdf", mime="application/pdf")
