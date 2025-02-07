import streamlit as st
import docx
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image
import os
from io import BytesIO

st.title("📄 PDF Merger Tool")
st.write("Upload multiple files, reorder them using the table below, and merge them into a single PDF.")

# Upload files
uploaded_files = st.file_uploader(
    "Upload Files",
    type=["txt", "docx", "pdf", "jpg", "png", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    # Create a DataFrame with two columns: Order and Filename.
    data = {"Order": list(range(1, len(uploaded_files) + 1)),
            "Filename": [f.name for f in uploaded_files]}
    df = pd.DataFrame(data)
    
    st.write("Reorder your files by editing the 'Order' column (lower numbers will appear first).")
    # Display a data editor so you can manually change the Order values.
    edited_df = st.data_editor("Reorder Files", df, num_rows="dynamic", use_container_width=True)

    if st.button("Merge Files"):
        # Sort the dataframe by the "Order" column.
        sorted_df = edited_df.sort_values("Order")
        ordered_filenames = sorted_df["Filename"].tolist()

        pdf_docs = []      # List to store PDF objects
        merged_text = ""   # String for TXT content

        for filename in ordered_filenames:
            # Find the file object with the matching filename.
            file_obj = next((f for f in uploaded_files if f.name == filename), None)
            if file_obj is None:
                continue

            ext = os.path.splitext(file_obj.name)[1].lower()

            if ext == ".txt":
                merged_text += file_obj.read().decode("utf-8") + "\n\n"
            elif ext == ".docx":
                doc = docx.Document(file_obj)
                doc_text = "\n".join([para.text for para in doc.paragraphs])
                # Convert DOCX text to PDF page
                text_pdf = fitz.open()
                text_page = text_pdf.new_page(width=595, height=842)
                text_rect = fitz.Rect(50, 50, 545, 800)
                text_page.insert_textbox(text_rect, doc_text, fontsize=12, fontname="helv")
                pdf_docs.append(text_pdf)
            elif ext == ".xlsx":
                df_excel = pd.read_excel(file_obj)
                excel_text = df_excel.to_string()
                # Convert XLSX text to PDF page
                text_pdf = fitz.open()
                text_page = text_pdf.new_page(width=595, height=842)
                text_rect = fitz.Rect(50, 50, 545, 800)
                text_page.insert_textbox(text_rect, excel_text, fontsize=12, fontname="helv")
                pdf_docs.append(text_pdf)
            elif ext == ".pdf":
                pdf_docs.append(fitz.open(stream=file_obj.read(), filetype="pdf"))
            elif ext in (".jpg", ".png"):
                img = Image.open(file_obj)
                img_bytes = BytesIO()
                img.save(img_bytes, format="PDF")
                pdf_docs.append(fitz.open("pdf", img_bytes.getvalue()))

        # If there is merged text (from TXT files), convert it to a PDF page.
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
