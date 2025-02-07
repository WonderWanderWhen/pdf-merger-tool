import streamlit as st
import docx
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image
import os
from io import BytesIO
from tabulate import tabulate  # For better formatting of Excel tables

st.title("📄 PDF Merger Tool")
st.write("Upload multiple files, specify the desired order using comma-separated indices, and merge them into a single PDF.")

# Upload files
uploaded_files = st.file_uploader(
    "Upload Files",
    type=["txt", "docx", "pdf", "jpg", "png", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    st.write("### Uploaded Files:")
    for i, f in enumerate(uploaded_files):
        st.write(f"{i+1}: {f.name}")

    order_input = st.text_input(
        "Enter desired order as comma-separated indices (e.g., 2,1,3). Leave blank to use the natural order."
    )

    if order_input:
        try:
            # Convert the input into a list of integers (1-indexed)
            order_list = [int(x.strip()) for x in order_input.split(",")]
            if len(order_list) != len(uploaded_files):
                st.error("The number of indices does not match the number of uploaded files.")
                ordered_files = uploaded_files
            else:
                # Convert to 0-indexed and reorder the files accordingly
                ordered_files = [uploaded_files[i - 1] for i in order_list]
        except Exception as e:
            st.error("Invalid input for order. Please enter comma-separated numbers.")
            ordered_files = uploaded_files
    else:
        ordered_files = uploaded_files

    if st.button("Merge Files"):
        pdf_docs = []     # List to store PDF objects for merging
        merged_text = ""  # For TXT file content

        for file_obj in ordered_files:
            ext = os.path.splitext(file_obj.name)[1].lower()

            if ext == ".txt":
                # Append TXT content to be merged on one page
                merged_text += file_obj.read().decode("utf-8") + "\n\n"

            elif ext == ".docx":
                # Read DOCX and convert its text to a PDF page
                doc = docx.Document(file_obj)
                doc_text = "\n".join([para.text for para in doc.paragraphs])
                text_pdf = fitz.open()
                text_page = text_pdf.new_page(width=595, height=842)
                text_rect = fitz.Rect(50, 50, 545, 800)
                text_page.insert_textbox(text_rect, doc_text, fontsize=12, fontname="helv")
                pdf_docs.append(text_pdf)

            elif ext == ".xlsx":
                # Read Excel and convert its content to a formatted table string
                df_excel = pd.read_excel(file_obj)
                # Use tabulate to format the table with grid lines
                excel_text = tabulate(df_excel, headers="keys", tablefmt="grid")
                text_pdf = fitz.open()
                text_page = text_pdf.new_page(width=595, height=842)
                text_rect = fitz.Rect(50, 50, 545, 800)
                # Use a slightly smaller font size to fit the table
                text_page.insert_textbox(text_rect, excel_text, fontsize=10, fontname="helv")
                pdf_docs.append(text_pdf)

            elif ext == ".pdf":
                # Open PDFs using PyMuPDF
                pdf_docs.append(fitz.open(stream=file_obj.read(), filetype="pdf"))

            elif ext in (".jpg", ".png"):
                # Convert images to a PDF page
                img = Image.open(file_obj)
                img_bytes = BytesIO()
                img.save(img_bytes, format="PDF")
                pdf_docs.append(fitz.open("pdf", img_bytes.getvalue()))

        # If there is merged text from TXT files, convert it to a PDF page and add it at the beginning.
        if merged_text:
            text_pdf = fitz.open()
            text_page = text_pdf.new_page(width=595, height=842)
            text_rect = fitz.Rect(50, 50, 545, 800)
            text_page.insert_textbox(text_rect, merged_text, fontsize=12, fontname="helv")
            pdf_docs.insert(0, text_pdf)

        if not pdf_docs:
            st.error("⚠️ No valid PDF, DOCX, XLSX, or image files to merge. Please upload valid files.")
        else:
            merged_doc = fitz.open()
            for pdf in pdf_docs:
                merged_doc.insert_pdf(pdf)

            output_path = "merged_output.pdf"
            merged_doc.save(output_path)

            st.success("✅ Merged PDF is ready!")
            st.download_button(
                "Download Merged PDF",
                open(output_path, "rb"),
                file_name="merged_output.pdf",
                mime="application/pdf"
            )
