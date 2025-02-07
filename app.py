import streamlit as st
import docx
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image
import os
from io import BytesIO
from tabulate import tabulate  # For formatting Excel tables

st.title("📄 PDF Merger Tool")
st.write(
    "Upload multiple files, then reorder them by editing the 'Order' column. "
    "After clicking 'Merge Files', the content of TXT, DOCX, XLSX, PDF, and image files "
    "will be merged into a single PDF."
)

# Upload files
uploaded_files = st.file_uploader(
    "Upload Files",
    type=["txt", "docx", "pdf", "jpg", "png", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    # Create a DataFrame with an 'Order' column and 'Filename'
    file_data = [{"Order": i + 1, "Filename": f.name} for i, f in enumerate(uploaded_files)]
    df = pd.DataFrame(file_data)
    
    st.write("### Reorder Files")
    st.write("Edit the 'Order' column (lower numbers appear first).")
    # Use st.data_editor without a positional label
    edited_df = st.data_editor(df, use_container_width=True)
    
    # Sort the DataFrame by the "Order" column to get the final order
    sorted_df = edited_df.sort_values("Order")
    ordered_filenames = sorted_df["Filename"].tolist()
    
    if st.button("Merge Files"):
        pdf_docs = []      # List to store PDF objects
        merged_text = ""   # For TXT file content
        
        # Process files in the specified order
        for filename in ordered_filenames:
            # Find the corresponding file object
            file_obj = next((f for f in uploaded_files if f.name == filename), None)
            if not file_obj:
                continue

            ext = os.path.splitext(file_obj.name)[1].lower()
            
            if ext == ".txt":
                # Read text file and append its content
                try:
                    merged_text += file_obj.read().decode("utf-8") + "\n\n"
                except Exception as e:
                    st.error(f"Error reading {file_obj.name}: {e}")
            elif ext == ".docx":
                # Read DOCX file, extract text, and convert to a PDF page
                try:
                    doc = docx.Document(file_obj)
                    doc_text = "\n".join([para.text for para in doc.paragraphs])
                    text_pdf = fitz.open()
                    text_page = text_pdf.new_page(width=595, height=842)
                    text_rect = fitz.Rect(50, 50, 545, 800)
                    text_page.insert_textbox(text_rect, doc_text, fontsize=12, fontname="helv")
                    pdf_docs.append(text_pdf)
                except Exception as e:
                    st.error(f"Error processing DOCX {file_obj.name}: {e}")
            elif ext == ".xlsx":
                # Read Excel file, format as grid using tabulate, and convert to PDF
                try:
                    df_excel = pd.read_excel(file_obj)
                    excel_text = tabulate(df_excel, headers="keys", tablefmt="grid")
                    text_pdf = fitz.open()
                    text_page = text_pdf.new_page(width=595, height=842)
                    # Use a larger text box area with smaller margins and font size 8
                    text_rect = fitz.Rect(20, 20, 575, 822)
                    text_page.insert_textbox(text_rect, excel_text, fontsize=8, fontname="helv")
                    pdf_docs.append(text_pdf)
                except Exception as e:
                    st.error(f"Error processing XLSX {file_obj.name}: {e}")
            elif ext == ".pdf":
                # Open PDF files
                try:
                    pdf_docs.append(fitz.open(stream=file_obj.read(), filetype="pdf"))
                except Exception as e:
                    st.error(f"Error processing PDF {file_obj.name}: {e}")
            elif ext in (".jpg", ".png"):
                # Convert image to PDF page
                try:
                    img = Image.open(file_obj)
                    img_bytes = BytesIO()
                    img.save(img_bytes, format="PDF")
                    pdf_docs.append(fitz.open("pdf", img_bytes.getvalue()))
                except Exception as e:
                    st.error(f"Error processing image {file_obj.name}: {e}")
        
        # If there's merged text from TXT files, convert it into a PDF page and insert it at the beginning.
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
