import streamlit as st
import docx
import fitz
import pandas as pd
from PIL import Image
import os
from io import BytesIO
import io
import matplotlib.pyplot as plt

st.title("📄 PDF Merger Tool")

uploaded_files = st.file_uploader(
    "Upload Files",
    type=["txt", "docx", "pdf", "jpg", "png", "xlsx"],
    accept_multiple_files=True
)

file_contents = []  # Store file contents here
if uploaded_files:
    for f in uploaded_files:
        file_contents.append({"name": f.name, "content": f.read(), "order": 0})  # Read content immediately

    st.write("### Reorder Files")

    for i, file_data in enumerate(file_contents):
        file_data["order"] = st.number_input(f"Order for {file_data['name']}", min_value=1, value=i + 1, step=1)

    if st.button("Merge Files"):
        file_contents.sort(key=lambda x: x["order"])  # Sort based on order

        pdf_docs = []
        merged_text = ""

        for file_data in file_contents:
            filename = file_data["name"]
            content = file_data["content"]
            ext = os.path.splitext(filename)[1].lower()

            try:
                if ext == ".txt":
                    merged_text += content.decode("utf-8") + "\n\n"

                elif ext == ".docx":
                    doc = docx.Document(io.BytesIO(content))  # Use BytesIO with content
                    doc_text = "\n".join([para.text for para in doc.paragraphs])
                    text_pdf = fitz.open()
                    text_page = text_pdf.new_page()
                    text_page.insert_textbox(fitz.Rect(50, 50, 550, 800), doc_text, fontsize=12)
                    pdf_docs.append(text_pdf)

                elif ext == ".xlsx":
                    df_excel = pd.read_excel(io.BytesIO(content))  # Use BytesIO

                    fig, ax = plt.subplots(figsize=(8, 6))
                    ax.table(cellText=df_excel.values, colLabels=df_excel.columns, loc='center')
                    ax.axis('off')
                    plt.tight_layout()

                    img_buf = io.BytesIO()
                    plt.savefig(img_buf, format='png', dpi=300)
                    img_buf.seek(0)

                    img = Image.open(img_buf).convert("RGB")
                    img_bytes = BytesIO()
                    img.save(img_bytes, format="PDF")
                    pdf_docs.append(fitz.open("pdf", img_bytes.getvalue()))
                    plt.close(fig)

                elif ext == ".pdf":
                    pdf_docs.append(fitz.open(stream=content, filetype="pdf"))  # Use content directly

                elif ext in (".jpg", ".png"):
                    img = Image.open(io.BytesIO(content)).convert("RGB")  # Use BytesIO
                    img_bytes = BytesIO()
                    img.save(img_bytes, format="PDF")
                    pdf_docs.append(fitz.open("pdf", img_bytes.getvalue()))

            except Exception as e:
                st.error(f"Error processing {filename}: {e}")
                st.stop()

        # ... (Rest of the merging and download code remains the same)
