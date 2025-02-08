import gradio as gr
import docx
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image
import os
from io import BytesIO
import re

# ✅ Function to merge files
def process_files(files, order, pdf_name):
    if not files or not order:
        return "Warning: No files uploaded or order not set!", None

    pdf_docs = []
    image_files = []
    merged_text = ""

    for index in order:
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

    if merged_text:
        text_pdf = fitz.open()
        text_page = text_pdf.new_page(width=595, height=842)  
        text_rect = fitz.Rect(50, 50, 545, 800)
        text_page.insert_textbox(text_rect, merged_text, fontsize=12, fontname="helv")
        pdf_docs.insert(0, text_pdf)

    merged_doc = fitz.open()

    for pdf in pdf_docs:
        merged_doc.insert_pdf(pdf)

    for img_file in image_files:
        img = Image.open(img_file)
        img_width, img_height = img.size
        a4_width, a4_height = 595, 842
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

    return f"PDF '{pdf_name}.pdf' is ready for download!", final_pdf_path

# ✅ Fix: Ensure file selection dropdown updates correctly
def reorder_files(files):
    if not files:
        return "Warning: Please upload files!", [], []

    file_names = [file.name for file in files]
    order_choices = list(range(len(file_names)))  # Ensure choices are properly set

    return gr.update(value=file_names, choices=file_names), order_choices

# ✅ Create Gradio interface
with gr.Blocks() as demo:
    gr.Markdown("# PDF Merger Tool")
    gr.Markdown("Upload files, set order, and get your merged PDF!")

    file_input = gr.Files(file_types=[".txt", ".docx", ".pdf", ".jpg", ".png", ".xlsx"], label="Upload Files", interactive=True)

    order_input = gr.Dropdown([], multiselect=True, label="Select Order (Drag to Rearrange)", interactive=True)
    
    pdf_name_input = gr.Textbox(label="Enter PDF Name", value="Merged_Document")

    submit_button = gr.Button("Merge & Download PDF")
    output_message = gr.Textbox(label="Status")
    output_pdf = gr.File(label="Download Merged PDF")

    # ✅ Fix: Ensure file selection dropdown updates properly
    file_input.change(reorder_files, [file_input], [order_input])

    submit_button.click(process_files, [file_input, order_input, pdf_name_input], [output_message, output_pdf])

# ✅ Launch Web App
demo.launch(share=True)
