import streamlit as st
import docx
import fitz  # PyMuPDF
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

if uploaded_files:
    file_data = [{"Order": i + 1, "Filename": f.name, "File": f} for i, f in enumerate(uploaded_files)]
    df = pd.DataFrame(file_data)

    st.write("### Reorder Files")
    st.write("Edit the 'Order' column (lower numbers will appear first).")
    edited_df = st.data_editor(df, use_container_width=True)

    sorted_df = edited_df.sort_values("Order")

    if st.button("Merge Files"):
        pdf_docs = []
        merged_text = ""

        for _, row in sorted_df.iterrows():
            file_obj = row['File']
            filename = row['Filename']
            ext = os.path.splitext(filename)[1].lower()

            try:
                if ext == ".txt":
                    merged_text += file_obj.read().decode("utf-8") + "\n\n"

                elif ext == ".docx":
                    doc = docx.Document(file_obj)
                    doc_text = "\n".join([para.text for para in doc.paragraphs])
                    text_pdf = fitz.open()
                    text_page = text_pdf.new_page()
                    text_page.insert_textbox(fitz.Rect(50, 50, 550, 800), doc_text, fontsize=12)
                    pdf_docs.append(text_pdf)

                elif ext == ".xlsx":
                    df_excel = pd.read_excel(file_obj)

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
                    pdf_docs.append(fitz.open(stream=file_obj.read(), filetype="pdf"))

                elif ext in (".jpg", ".png"):
                    img = Image.open(file_obj).convert("RGB")
                    img_bytes = BytesIO()
                    img.save(img_bytes, format="PDF")
                    pdf_docs.append(fitz.open("pdf", img_bytes.getvalue()))

            except Exception as e:
                st.error(f"Error processing {filename}: {e}")
                st.stop()

        if merged_text:
            text_pdf = fitz.open()
            text_page = text_pdf.new_page()
            text_page.insert_textbox(fitz.Rect(50, 50, 550, 800), merged_text, fontsize=12)
            pdf_docs.insert(0, text_pdf)

        if not pdf_docs:
            st.error("⚠️ No valid files to merge.")
        else:
            merged_doc = fitz.open()
            for pdf in pdf_docs:
                merged_doc.insert_pdf(pdf)

            output_path = "merged_output.pdf"
            merged_doc.save(output_path)

            st.success("✅ Merged PDF is ready!")
            with open(output_path, "rb") as f:
                st.download_button(
                    "Download Merged PDF",
                    data=f,
                    file_name="merged_output.pdf",
                    mime="application/pdf"
                )
