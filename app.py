import gradio as gr
import docx
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image
import os
from io import BytesIO
import tempfile

def ensure_dict(file_item):
    """
    Convert an uploaded file into a simple dictionary with keys "name" and "data".
    If it's already a dict, return it as is.
    """
    if isinstance(file_item, dict):
        return file_item
    try:
        data = file_item.read()
    except Exception as e:
        data = None
    return {"name": file_item.name, "data": data}

def process_files(files, order, pdf_name):
    if not files or not order:
        return "Warning: No files uploaded or order not set!", None

    # Force every uploaded file into a simple dict.
    files = [ensure_dict(f) for f in files]
    
    pdf_docs = []      # List to store PDFs (or text converted to PDF)
    image_files = []   # List to store image file paths (temporary)
    merged_text = ""   # To accumulate text content

    # Process files in the order specified by the dropdown.
    # Here, each item in "order" is a string like "0: filename"
    for o in order:
        try:
            idx = int(o.split(":")[0])
        except Exception as e:
            print("Error converting order element", o, e)
            continue

        file_item = files[idx]
        fname = file_item.get("name", "uploaded_file")
        file_data = file_item.get("data")
        # Write the file data to a temporary file.
        temp_path = os.path.join(tempfile.gettempdir(), fname)
        with open(temp_path, "wb") as f:
            f.write(file_data)
        filename = temp_path
        ext = os.path.splitext(filename)[1].lower()

        if ext == ".txt":
            with open(filename, "r", encoding="utf-8") as f:
                merged_text += f.read() + "\n\n"
        elif ext == ".docx":
            doc = docx.Document(filename)
            merged_text += "\n".join([para.text for para in doc.paragraphs]) + "\n\n"
        elif ext == ".pdf":
            pdf_docs.append(fitz.open(filename))
        elif ext in (".jpg", ".png"):
            image_files.append(filename)
        elif ext == ".xlsx":
            df = pd.read_excel(filename)
            merged_text += df.to_string() + "\n\n"

    # If any text was accumulated, convert it into a PDF page.
    if merged_text:
        text_pdf = fitz.open()
        text_page = text_pdf.new_page(width=595, height=842)  # A4 page
        text_rect = fitz.Rect(50, 50, 545, 800)
        text_page.insert_textbox(text_rect, merged_text, fontsize=12, fontname="helv")
        pdf_docs.insert(0, text_pdf)

    # Create a new PDF document and merge all PDF pages.
    merged_doc = fitz.open()
    for pdf in pdf_docs:
        merged_doc.insert_pdf(pdf)

    # Process each image: resize to fit an A4 page and add as a full page.
    for img_file in image_files:
        img = Image.open(img_file)
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

    final_pdf_path = f"{pdf_name}.pdf"
    merged_doc.save(final_pdf_path)
    
    # Return the merged PDF as a dict so that gr.File can handle it.
    with open(final_pdf_path, "rb") as f:
        pdf_bytes = f.read()
    return f"PDF '{pdf_name}.pdf' is ready for download!", {"name": f"{pdf_name}.pdf", "data": pdf_bytes}

def reorder_files(files):
    """
    Build a list of simple strings for each file in the form "index: filename".
    This avoids passing complex objects to the dropdown.
    """
    if not files:
        return gr.update(value=[], choices=[]), []
    # Force conversion to dict for each file.
    files = [ensure_dict(f) for f in files]
    choices = [f"{i}: {str(files[i].get('name', ''))}" for i in range(len(files))]
    return gr.update(value=choices, choices=choices), choices

with gr.Blocks() as demo:
    gr.Markdown("# PDF Merger Tool")
    gr.Markdown("Upload files, arrange them via drag-and-drop, and get your merged PDF!")
    
    # Use gr.Files with type="bytes" to get plain dicts.
    file_input = gr.Files(
        file_types=[".txt", ".docx", ".pdf", ".jpg", ".png", ".xlsx"],
        label="Upload Files",
        interactive=True,
        type="bytes"
    )
    order_input = gr.Dropdown([], multiselect=True, label="Select Order (Drag to Rearrange)", interactive=True)
    pdf_name_input = gr.Textbox(label="Enter PDF Name", value="Merged_Document")
    submit_button = gr.Button("Merge & Download PDF")
    output_message = gr.Textbox(label="Status")
    output_pdf = gr.File(label="Download Merged PDF")
    
    file_input.change(reorder_files, [file_input], [order_input])
    submit_button.click(process_files, [file_input, order_input, pdf_name_input], [output_message, output_pdf])

demo.launch(share=True)
