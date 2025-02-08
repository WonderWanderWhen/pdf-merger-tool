import gradio as gr
import docx
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image
import os
from io import BytesIO
import re

# Function to merge files into one PDF
def process_files(files, order, pdf_name):
    if not files or not order:
        return "Warning: No files uploaded or order not set!", None

    pdf_docs = []      # List to store PDF documents (or text converted to PDF)
    image_files = []   # List to store image file paths
    merged_text = ""   # String to accumulate text content

    # Process files in the user-specified order
    for index in order:
        filename = files[index].name
        ext = os.path.splitext(filename)[1].lower()

        if ext == ".txt":  # Process plain text files
            with open(filename, "r", encoding="utf-8") as f:
                merged_text += f.read() + "\n\n"

        elif ext == ".docx":  # Process Word documents
            doc = docx.Document(filename)
            merged_text += "\n".join([para.text for para in doc.paragraphs]) + "\n\n"

        elif ext == ".pdf":  # For PDFs, open them with PyMuPDF
            pdf_docs.append(fitz.open(filename))

        elif ext in (".jpg", ".png"):  # For images, store the file path
            image_files.append(filename)

        elif ext == ".xlsx":  # For Excel files, convert data to text
            df = pd.read_excel(filename)
            merged_text += df.to_string() + "\n\n"

    # Convert the merged text to a selectable PDF page (if any text exists)
    if merged_text:
        text_pdf = fitz.open()
        text_page = text_pdf.new_page(width=595, height=842)  # A4 page size
        text_rect = fitz.Rect(50, 50, 545, 800)  # Set margins
        text_page.insert_textbox(text_rect, merged_text, fontsize=12, fontname="helv")
        pdf_docs.insert(0, text_pdf)  # Add text PDF at the beginning

    # Create a new PDF document to merge all pages
    merged_doc = fitz.open()
    for pdf in pdf_docs:
        merged_doc.insert_pdf(pdf)

    # Process and add each image as a full page in the PDF
    for img_file in image_files:
        img = Image.open(img_file)
        img_width, img_height = img.size
        a4_width, a4_height = 595, 842  # A4 page dimensions
        scale = min(a4_width / img_width, a4_height / img_height)
        new_size = (int(img_width * scale), int(img_height * scale))
        img = img.resize(new_size)
        
        # Create a new PDF page for the image
        img_pdf = fitz.open()
        img_page = img_pdf.new_page(width=a4_width, height=a4_height)
        img_bytes = BytesIO()
        img.save(img_bytes, format="JPEG")
        img_page.insert_image(fitz.Rect(0, 0, new_size[0], new_size[1]), stream=img_bytes.getvalue())
        merged_doc.insert_pdf(img_pdf)

    # Save the final merged PDF with the custom name
    final_pdf_path = f"{pdf_name}.pdf"
    merged_doc.save(final_pdf_path)

    return f"PDF '{pdf_name}.pdf' is ready for download!", final_pdf_path

# Updated function to extract file names as strings for reordering
def reorder_files(files):
    if not files:
        return gr.update(value=[], choices=[]), []
    file_names = []
    for file in files:
        try:
            name = file.name
        except AttributeError:
            name = file.get("name", "")
        file_names.append(str(name))
    order_choices = list(range(len(file_names)))
    return gr.update(value=file_names, choices=file_names), order_choices

# Create the Gradio interface
with gr.Blocks() as demo:
    gr.Markdown("# PDF Merger Tool")
    gr.Markdown("Upload files, set order (via drag-and-drop), and get your merged PDF!")

    # File uploader allowing multiple files (using gr.Files)
    file_input = gr.Files(
        file_types=[".txt", ".docx", ".pdf", ".jpg", ".png", ".xlsx"],
        label="Upload Files",
        interactive=True
    )

    # Dropdown to set the order; updated by the reorder_files function
    order_input = gr.Dropdown([], multiselect=True, label="Select Order (Drag to Rearrange)", interactive=True)
    
    # Input for custom PDF name
    pdf_name_input = gr.Textbox(label="Enter PDF Name", value="Merged_Document")

    # Button to trigger the merging process
    submit_button = gr.Button("Merge & Download PDF")

    # Outputs: a status message and a file download component
    output_message = gr.Textbox(label="Status")
    output_pdf = gr.File(label="Download Merged PDF")

    # When files are uploaded, update the order dropdown
    file_input.change(reorder_files, [file_input], [order_input])

    # On button click, process the files and produce the merged PDF
    submit_button.click(process_files, [file_input, order_input, pdf_name_input], [output_message, output_pdf])

# Launch the app with share=True for a public URL
demo.launch(share=True)
