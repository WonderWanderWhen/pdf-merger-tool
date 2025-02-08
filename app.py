import gradio as gr
import docx
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image
import os
from io import BytesIO

def process_files(files, order, pdf_name):
    """
    Process the uploaded files (which are file paths) in the user-specified order,
    merge their contents into a single PDF, and return it.
    """
    if not files or not order:
        return "Warning: No files uploaded or order not set!", None

    pdf_docs = []      # List to store PDFs (or text converted to PDF)
    image_files = []   # List to store image file paths
    merged_text = ""   # To accumulate text content

    # 'order' is a list of strings like "0: filename". Extract the index.
    for o in order:
        try:
            idx = int(o.split(":")[0])
        except Exception as e:
            print("Error converting order element", o, e)
            continue

        file_path = files[idx]  # file_path is a string
        fname = os.path.basename(file_path)
        ext = os.path.splitext(fname)[1].lower()

        if ext == ".txt":
            with open(file_path, "r", encoding="utf-8") as f:
                merged_text += f.read() + "\n\n"
        elif ext == ".docx":
            doc = docx.Document(file_path)
            merged_text += "\n".join([para.text for para in doc.paragraphs]) + "\n\n"
        elif ext == ".pdf":
            pdf_docs.append(fitz.open(file_path))
        elif ext in (".jpg", ".png"):
            image_files.append(file_path)
        elif ext == ".xlsx":
            df = pd.read_excel(file_path)
            merged_text += df.to_string() + "\n\n"

    # Convert any merged text into a PDF page.
    if merged_text:
        text_pdf = fitz.open()
        text_page = text_pdf.new_page(width=595, height=842)  # A4 size
        text_rect = fitz.Rect(50, 50, 545, 800)
        text_page.insert_textbox(text_rect, merged_text, fontsize=12, fontname="helv")
        pdf_docs.insert(0, text_pdf)  # Insert text PDF at the beginning

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
    
    # Read the merged PDF as bytes and return a dict (so gr.File can handle it)
    with open(final_pdf_path, "rb") as f:
        pdf_bytes = f.read()
    return f"PDF '{pdf_name}.pdf' is ready for download!", {"name": f"{pdf_name}.pdf", "data": pdf_bytes}

def reorder_files(files):
    """
    Build a list of simple strings (e.g., "0: filename") for the ordering dropdown.
    Since files are now plain strings (file paths), we extract the basename.
    """
    if not files:
        return gr.update(value=[], choices=[]), []
    choices = []
    for i, file_path in enumerate(files):
        fname = os.path.basename(file_path)
        choices.append(f"{i}: {fname}")
    return gr.update(value=choices, choices=choices), choices

with gr.Blocks() as demo:
    gr.Markdown("# PDF Merger Tool")
    gr.Markdown("Upload files, arrange them via drag-and-drop, and get your merged PDF!")
    
    # Use gr.Files with type="file" to get file paths (plain strings)
    file_input = gr.Files(
        file_types=[".txt", ".docx", ".pdf", ".jpg", ".png", ".xlsx"],
        label="Upload Files",
        interactive=True,
        type="file"
    )
    order_input = gr.Dropdown([], multiselect=True, label="Select Order (Drag to Rearrange)", interactive=True)
    pdf_name_input = gr.Textbox(label="Enter PDF Name", value="Merged_Document")
    submit_button = gr.Button("Merge & Download PDF")
    output_message = gr.Textbox(label="Status")
    output_pdf = gr.File(label="Download Merged PDF")
    
    # Update the ordering dropdown when files are uploaded.
    file_input.change(reorder_files, [file_input], [order_input])
    # Process files on button click.
    submit_button.click(process_files, [file_input, order_input, pdf_name_input], [output_message, output_pdf])
    
demo.launch(share=True)
