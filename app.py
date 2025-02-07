import streamlit as st
import docx
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image
import os
from io import BytesIO
from tabulate import tabulate

def process_uploaded_files(uploaded_files):
    """Process uploaded files and return a list of PDF documents"""
    pdf_docs = []
    merged_text = ""

    for file_obj in uploaded_files:
        ext = os.path.splitext(file_obj.name)[1].lower()
        try:
            if ext == ".txt":
                merged_text += file_obj.read().decode("utf-8") + "\n\n"
            elif ext == ".docx":
                doc = docx.Document(file_obj)
                doc_text = "\n".join([para.text for para in doc.paragraphs])
                text_pdf = fitz.open()
                text_page = text_pdf.new_page(width=595, height=842)
                text_rect = fitz.Rect(50, 50, 545, 800)
                text_page.insert_textbox(text_rect, doc_text, fontsize=12, fontname="helv")
                pdf_docs.append(text_pdf)
            elif ext == ".xlsx":
                df_excel = pd.read_excel(file_obj)
                excel_text = tabulate(df_excel, headers="keys", tablefmt="grid")
                text_pdf = fitz.open()
                text_page = text_pdf.new_page(width=595, height=842)
                text_rect = fitz.Rect(20, 20, 575, 822)
                text_page.insert_textbox(text_rect, excel_text, fontsize=6, fontname="helv")
                pdf_docs.append(text_pdf)
            elif ext == ".pdf":
                pdf_docs.append(fitz.open(stream=file_obj.read(), filetype="pdf"))
            elif ext in (".jpg", ".png"):
                img = Image.open(file_obj)
                img_bytes = BytesIO()
                img.save(img_bytes, format="PDF")
                pdf_docs.append(fitz.open("pdf", img_bytes.getvalue()))
        except Exception as e:
            st.error(f"Error processing {file_obj.name}: {str(e)}")

    if merged_text:
        text_pdf = fitz.open()
        text_page = text_pdf.new_page(width=595, height=842)
        text_rect = fitz.Rect(50, 50, 545, 800)
        text_page.insert_textbox(text_rect, merged_text, fontsize=12, fontname="helv")
        pdf_docs.insert(0, text_pdf)

    return pdf_docs

def create_merged_pdf(pdf_docs):
    """Create a merged PDF document from the list of PDF documents"""
    merged_doc = fitz.open()
    for pdf in pdf_docs:
        merged_doc.insert_pdf(pdf)
    return merged_doc

def download_merged_pdf(merged_doc):
    """Download the merged PDF document"""
    output_buffer = BytesIO()
    merged_doc.save(output_buffer)
    merged_doc.close()
    st.success(" Merged PDF is ready!")
    st.download_button(
        "Download Merged PDF",
        data=output_buffer.getvalue(),
        file_name="merged_output.pdf",
        mime="application/pdf"
    )

st.title(" PDF Merger Tool")
st.write("Upload multiple files and specify the desired order using comma-separated indices (e.g., 2,1,3).")

uploaded_files = st.file_uploader(
    "Upload Files",
    type=["txt", "docx", "pdf", "jpg", "png", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    st.write("### Uploaded Files:")
    # Display only file names (strings) to avoid Arrow issues
    file_names = [file.name for file in uploaded_files]
    for idx, name in
