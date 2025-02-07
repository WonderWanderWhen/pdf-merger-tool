import streamlit as st
import docx
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image
import os
from io import BytesIO
from tabulate import tabulate  # For formatting Excel tables

st.title("📄 PDF Merger Tool")
st.write("Upload multiple files and specify the desired order using comma-separated indices (e.g., 2,1,3).")

# Upload files
uploaded_files = st.file_uploader(
    "Upload Files",
    type=["txt", "docx", "pdf", "jpg", "png", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    st.write("### Uploaded Files:")
    # List the uploaded files with indices (using st.text to avoid Arrow conversion issues)
    for idx, file in enumerate(uploaded_files):
        st.text(f"{idx+1}: {file.name}")

    # Default order is the natural order (e.g., "1,2,3,...")
    default_order = ",".join([str(i + 1) for i in range(len(uploaded_files))])
    order_input = st.text_input(
        "Enter desired order as comma-separated indices (e.g., 2,1,3)", value=default_order
    )

    if st.button("Merge Files"):
        try:
            # Convert the input string into a list of integers (1-indexed)
            order_list = [int(x.strip()) for x in order_input.split(",") if x.strip() != ""]
            if len(order_list) != len(uploaded_files):
                st.error("The number of indices does not match the number of uploaded files.")
                ordered_files = uploaded_files
            else:
                if any(i < 1 or i > len(uploaded_files) for i in order_list):
                    st.error("One or more indices are out of range.")
                    ordered_files = uploaded_files
                else:
                    # Reorder files using the provided indices (convert to 0-indexed)
                    ordered_files = [uploaded_files[i - 1] for i in order_list]
        except Exception as e:
            st.error("Invalid input for order. Please enter comma-separated numbers.")
            ordered_files = uploaded_files

        pdf_docs = []     # List to store PDF objects for merging
        merged_text = ""  # For TXT file content

        # Process each file in the specified order
        for file_obj in ordered_files:
            ext = os.path.splitext(file_obj.name)[1].lower()
            if ext == ".txt":
                try:
                    merged_text += file_obj.read().decode("utf-8") + "\n\n"
                except Exception as e:
                    st.error(f"Error reading {file_obj.name}: {e}")
            elif ext == ".docx":
                try:
                    doc = docx.Document(file_obj)
                    doc_text = "\n".join([para.text for para in doc.paragraphs])
                    # Convert DOCX text to a PDF page
                    text_pdf = fitz.open()
                    text_page = text_pdf.new_page(width=595, height=842)
                    text_rect = fitz.Rect(50, 50, 545, 800)
                    text_page.insert_textbox(text_rect, doc_text, fontsize=12, fontname="helv")
                    pdf_docs.append(text_pdf)
                except Exception as e:
                    st.error(f"Error processing DOCX {file_obj.name}: {e}")
            elif ext == ".xlsx":
                try:
                    file_obj.seek(0)
                    df_excel = pd.read_excel(file_obj)
                    # Format the DataFrame as a grid with headers using tabulate
                    excel_text = tabulate(df_excel, headers="keys", tablefmt="grid")
                    text_pdf = fitz.open()
                    text_page = text_pdf.new_page(width=595, height=842)
                    # Use nearly full page dimensions and a small font size to show more content
                    text_rect = fitz.Rect(20, 20, 575, 822)
                    text_page.insert_textbox(text_rect, excel_text, fontsize=6, fontname="helv")
                    pdf_docs.append(text_pdf)
                except Exception as e:
                    st.error(f"Error processing XLSX {file_obj.name}: {e}")
            elif ext == ".pdf":
                try:
                    pdf_docs.append(fitz.open(stream=file_obj.read(), filetype="pdf"))
                except Exception as e:
                    st.error(f"Error processing PDF {file_obj.name}: {e}")
            elif ext in (".jpg", ".png"):
                try:
                    img = Image.open(file_obj)
                    img_bytes = BytesIO()
                    img.save(img_bytes, format="PDF")
                    pdf_docs.append(fitz.open("pdf", img_bytes.getvalue()))
                except Exception as e:
                    st.error(f"Error processing image {file_obj.name}: {e}")

        # If there is merged text from TXT files, convert it into a PDF page and insert at the beginning.
        if merged_text:
            text_pdf = fitz.open()
            text_page = text_pdf.new_page(width=595, height=842)
            text_rect = fitz.Rect(50, 50, 545, 800)
            text_page.insert_textbox(text_rect, merged_text, fontsize=12, fontname="helv")
            pdf_docs.insert(0, text_pdf)

        if not pdf_docs:
            st.error("⚠️ No valid files to merge. Please upload valid files.")
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
