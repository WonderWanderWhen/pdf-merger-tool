import streamlit as st
import docx
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image
import os
from io import BytesIO

def merge_files(file_list, order_list, pdf_name):
    """
    Merge uploaded files (using their bytes from getvalue()) in the specified order.
    file_list: List of UploadedFile objects from st.file_uploader.
    order_list: List of integer indices representing the desired order.
    pdf_name: Name for the merged PDF.
    Returns: The merged PDF as bytes.
    """
    pdf_docs = []    # List to hold PyMuPDF document objects (or text turned into PDF)
    image_files = [] # List to hold UploadedFile objects for images
    merged_text = "" # To accumulate text content from TXT, DOCX, XLSX files

    for idx in order_list:
        try:
            file_obj = file_list[idx]
        except IndexError:
            st.error(f"Index {idx} is out of range for the uploaded files.")
            continue

        fname = file_obj.name
        ext = os.path.splitext(fname)[1].lower()

        if ext == ".txt":
            try:
                # Decode the text bytes
                text = file_obj.getvalue().decode("utf-8")
                merged_text += text + "\n\n"
            except Exception as e:
                st.error(f"Error reading {fname}: {e}")
        elif ext == ".docx":
            try:
                # Wrap the bytes in a BytesIO and read with python-docx
                doc = docx.Document(BytesIO(file_obj.getvalue()))
                text = "\n".join([para.text for para in doc.paragraphs])
                merged_text += text + "\n\n"
            except Exception as e:
                st.error(f"Error processing DOCX {fname}: {e}")
        elif ext == ".pdf":
            try:
                # Open PDF from bytes stream; specify filetype so PyMuPDF knows it's a PDF
                pdf = fitz.open(stream=file_obj.getvalue(), filetype="pdf")
                pdf_docs.append(pdf)
            except Exception as e:
                st.error(f"Error processing PDF {fname}: {e}")
        elif ext in (".jpg", ".png"):
            # For images, we'll process later; just store the file object.
            image_files.append(file_obj)
        elif ext == ".xlsx":
            try:
                # Use BytesIO so that pandas reads from the file bytes
                df = pd.read_excel(BytesIO(file_obj.getvalue()))
                merged_text += df.to_string() + "\n\n"
            except Exception as e:
                st.error(f"Error processing Excel file {fname}: {e}")

    # If any text content was collected, convert it into a PDF page.
    if merged_text:
        text_pdf = fitz.open()
        text_page = text_pdf.new_page(width=595, height=842)  # A4 page size in points
        text_rect = fitz.Rect(50, 50, 545, 800)  # Set margins
        text_page.insert_textbox(text_rect, merged_text, fontsize=12, fontname="helv")
        pdf_docs.insert(0, text_pdf)  # Prepend the text page

    # Create a new PDF document and merge all PDF pages.
    merged_doc = fitz.open()
    for pdf in pdf_docs:
        merged_doc.insert_pdf(pdf)

    # Process each image: open from BytesIO, resize to fit an A4 page, and add as a full page.
    for img_obj in image_files:
        try:
            im = Image.open(BytesIO(img_obj.getvalue()))
            im_width, im_height = im.size
            a4_width, a4_height = 595, 842  # A4 dimensions in points
            scale = min(a4_width / im_width, a4_height / im_height)
            new_size = (int(im_width * scale), int(im_height * scale))
            im = im.resize(new_size)
            img_pdf = fitz.open()
            img_page = img_pdf.new_page(width=a4_width, height=a4_height)
            img_bytes = BytesIO()
            im.save(img_bytes, format="JPEG")
            img_page.insert_image(fitz.Rect(0, 0, new_size[0], new_size[1]), stream=img_bytes.getvalue())
            merged_doc.insert_pdf(img_pdf)
        except Exception as e:
            st.error(f"Error processing image {img_obj.name}: {e}")

    # Save the merged PDF to a BytesIO object.
    output = BytesIO()
    merged_doc.save(output)
    output.seek(0)
    return output.getvalue()

# Streamlit UI
st.title("PDF Merger Tool")

st.write("Upload files (TXT, DOCX, PDF, JPG, PNG, XLSX) and specify the order as comma-separated indices.")

# Upload multiple files.
uploaded_files = st.file_uploader("Upload Files", accept_multiple_files=True)

if uploaded_files:
    st.write("Uploaded Files:")
    for i, f in enumerate(uploaded_files):
        st.write(f"{i}: {f.name}")

# Text input for order; user enters comma-separated indices.
order_input = st.text_input("Enter file order as comma-separated indices (e.g. '0,2,1')", value="")

# Input for PDF name.
pdf_name = st.text_input("Enter name for merged PDF (without .pdf extension)", value="Merged_Document")

if st.button("Merge PDF"):
    if not uploaded_files:
        st.error("Please upload at least one file.")
    else:
        try:
            if order_input.strip() == "":
                order_list = list(range(len(uploaded_files)))
            else:
                order_list = [int(x.strip()) for x in order_input.split(",")]
        except Exception as e:
            st.error("Invalid order input. Please enter comma-separated indices.")
            order_list = list(range(len(uploaded_files)))
        
        with st.spinner("Merging files..."):
            merged_pdf_bytes = merge_files(uploaded_files, order_list, pdf_name)
        st.success("Merged PDF is ready!")
        st.download_button("Download Merged PDF", data=merged_pdf_bytes, file_name=f"{pdf_name}.pdf", mime="application/pdf")
