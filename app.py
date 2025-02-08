import streamlit as st
import docx
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image
import os
from io import BytesIO
import tempfile

def merge_files(file_list, order_list, pdf_name):
    """
    Merge uploaded files into a single PDF according to the order_list.
    file_list: List of UploadedFile objects from st.file_uploader.
    order_list: List of integer indices representing the desired order.
    pdf_name: Name for the merged PDF.
    Returns: The merged PDF as bytes.
    """
    pdf_docs = []    # Will hold PyMuPDF document objects (or a PDF generated from text)
    image_files = [] # Will hold file paths for images
    merged_text = "" # Accumulate text from TXT, DOCX, XLSX files

    # Process each file in the specified order.
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
                text = file_obj.read().decode("utf-8")
                merged_text += text + "\n\n"
            except Exception as e:
                st.error(f"Error reading {fname}: {e}")
        elif ext == ".docx":
            # For DOCX, write the bytes to a temporary file and use python-docx.
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                tmp.write(file_obj.getvalue())
                tmp_path = tmp.name
            try:
                doc = docx.Document(tmp_path)
                text = "\n".join([para.text for para in doc.paragraphs])
                merged_text += text + "\n\n"
            except Exception as e:
                st.error(f"Error processing {fname}: {e}")
            finally:
                os.remove(tmp_path)
        elif ext == ".pdf":
            # Write PDF bytes to a temporary file and open with PyMuPDF.
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                tmp.write(file_obj.getvalue())
                tmp_path = tmp.name
            try:
                pdf = fitz.open(tmp_path)
                pdf_docs.append(pdf)
            except Exception as e:
                st.error(f"Error opening PDF {fname}: {e}")
            finally:
                os.remove(tmp_path)
        elif ext in (".jpg", ".png"):
            # For images, save the image bytes to a temporary file.
            with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                tmp.write(file_obj.getvalue())
                tmp_path = tmp.name
            image_files.append(tmp_path)
        elif ext == ".xlsx":
            try:
                df = pd.read_excel(file_obj)
                merged_text += df.to_string() + "\n\n"
            except Exception as e:
                st.error(f"Error processing Excel file {fname}: {e}")

    # If any text content was collected, convert it into a PDF page.
    if merged_text:
        text_pdf = fitz.open()
        # Create a new A4 page (595 x 842 points)
        text_page = text_pdf.new_page(width=595, height=842)
        text_rect = fitz.Rect(50, 50, 545, 800)
        text_page.insert_textbox(text_rect, merged_text, fontsize=12, fontname="helv")
        pdf_docs.insert(0, text_pdf)

    # Create a new PyMuPDF document to merge all pages.
    merged_doc = fitz.open()
    for pdf in pdf_docs:
        merged_doc.insert_pdf(pdf)

    # Process each image: resize to fit an A4 page and add as a page.
    for img_path in image_files:
        try:
            img = Image.open(img_path)
            img_width, img_height = img.size
            a4_width, a4_height = 595, 842  # A4 dimensions in points
            scale = min(a4_width / img_width, a4_height / img_height)
            new_size = (int(img_width * scale), int(img_height * scale))
            img = img.resize(new_size)
            img_pdf = fitz.open()
            img_page = img_pdf.new_page(width=a4_width, height=a4_height)
            img_bytes = BytesIO()
            img.save(img_bytes, format="JPEG")
            img_page.insert_image(fitz.Rect(0, 0, new_size[0], new_size[1]), stream=img_bytes.getvalue())
            merged_doc.insert_pdf(img_pdf)
        except Exception as e:
            st.error(f"Error processing image {img_path}: {e}")
        finally:
            if os.path.exists(img_path):
                os.remove(img_path)

    # Save the merged PDF to a temporary file.
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        merged_doc.save(tmp.name)
        merged_pdf_path = tmp.name

    # Read the merged PDF as bytes.
    with open(merged_pdf_path, "rb") as f:
        merged_pdf_bytes = f.read()
    os.remove(merged_pdf_path)
    return merged_pdf_bytes

# Streamlit UI
st.title("PDF Merger Tool")

st.write("Upload files (TXT, DOCX, PDF, JPG, PNG, XLSX) and specify the order (comma-separated indices).")

# Upload multiple files.
uploaded_files = st.file_uploader("Upload Files", accept_multiple_files=True)
# Show the list of uploaded file names with their indices.
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
                # If no order is provided, use natural order.
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
