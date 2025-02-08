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
    # If no files or no order selected, return a warning.
    if not files or not order:
        return "Warning: No files uploaded or order not set!", None

    # Convert the order (list of strings like "0: filename") into integer indices.
    indices = []
    for o in order:
        try:
            # Expecting a string like "0: some_filename.docx"
            idx = int(o.split(":")[0])
            indices.append(idx)
        except Exception as e:
            print("Error converting order element", o, e)

    pdf_docs = []      # List to store PDFs (or text converted to PDF)
    image_files = []   # List to store image file paths
    merged_text = ""   # Accumulate text from TXT, DOCX, XLSX

    # Process files in the user-specified order.
    for index in indices:
        # Use the file object from the files list.
        filename = files[index].name
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

    # Convert any merged text into a selectable PDF page.
    if merged_text:
        text_pdf = fitz.open()
        text_page = text_pdf.new_page(width=595, height=842)  # A4 page size
        text_rect = fitz.Rect(50, 50, 545, 800)  # Margins
        text_page.insert_textbox(text_rect, merged_text, fontsize=12, fontname="helv")
        pdf_docs.insert(0, text_pdf)  # Place text as the first page

    # Create a new PDF document to merge all pages.
    merged_doc = fitz.open()
    for pdf in pdf_docs:
        merged_doc.insert_pdf(pdf)

    # Process and add each image as a full page.
    for img_file in image_files:
        img = Image.open(img_file)
        img_width, img_height = img.size
        a4_width, a4_height = 595, 842  # A4 dimensions in points
        scale = min(a4_width / img_width, a4_height / img_height)
        new_size = (int(img_width * scale), int(img_height * scale))
        img = img.resize(new_size)

        # Create a new PDF page for the image.
        img_pdf = fitz.open()
        img_page = img_pdf.new_page(width=a4_width, height=a4_height)
        img_bytes = BytesIO()
        img.save(img_bytes, format="JPEG")
        img_page.insert_image(fitz.Rect(0, 0, new_size[0], new_size[1]), stream=img_bytes.getvalue())
        merged_doc.insert_pdf(img_pdf)

    # Save the merged PDF with the custom name.
    final_pdf_path = f"{pdf_name}.pdf"
    merged_doc.save(final_pdf_path)

    return f"PDF '{pdf_name}.pdf' is ready for download!", final_pdf_path

# Updated function to populate the ordering dropdown.
def reorder_files(files):
    if not files:
        return gr.update(value=[], choices=[]), []
    # Build a list of choices as strings "index: filename".
    choices = []
    for i, file in enumerate(files):
        try:
            name = file.name
        except Exception as e:
            name = str(file)
        choices.append(f"{i}: {name}")
    # Return the update along with the raw choices list.
    return gr.update(value=choices, choices=choices), choices

# Create the Gradio interface.
with gr.Blocks() as demo:
    gr.Markdown("# PDF Merger Tool")
    gr.Markdown("Upload files, arrange them via drag-and-drop, and get your merged PDF!")

    # File uploader allowing multiple files.
    file_input = gr.Files(
        file_types=[".txt", ".docx", ".pdf", ".jpg", ".png", ".xlsx"],
        label="Upload Files",
        interactive=True
    )
    # Dropdown to select order; it will be populated automatically.
    order_input = gr.Dropdown([], multiselect=True, label="Select Order (Drag to Rearrange)", interactive=True)
    # Input for a custom PDF name.
    pdf_name_input = gr.Textbox(label="Enter PDF Name", value="Merged_Document")
    # Button to trigger the merging.
    submit_button = gr.Button("Merge & Download PDF")
    # Outputs: a status message and a file download component.
    output_message = gr.Textbox(label="Status")
    output_pdf = gr.File(label="Download Merged PDF")

    # Update the dropdown when files are uploaded.
    file_input.change(reorder_files, [file_input], [order_input])
    # Process files on button click.
    submit_button.click(process_files, [file_input, order_input, pdf_name_input], [output_message, output_pdf])

# Launch the app with share=True to generate a public URL.
demo.launch(share=True)
