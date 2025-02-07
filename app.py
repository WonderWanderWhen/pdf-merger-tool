import streamlit as st
import docx
import fitz
import pandas as pd
from PIL import Image
import os
from io import BytesIO
import io
import matplotlib.pyplot as plt

def process_file(uploaded_file):
    filename = uploaded_file.name
    ext = os.path.splitext(filename)[1].lower()
    
    try:
        content = uploaded_file.getvalue()
        
        if ext == ".txt":
            return {"type": "text", "content": content.decode("utf-8")}
        
        elif ext == ".docx":
            doc = docx.Document(io.BytesIO(content))
            doc_text = "\n".join([para.text for para in doc.paragraphs])
            return {"type": "docx", "content": doc_text}
        
        elif ext == ".xlsx":
            return {"type": "xlsx", "content": content}
        
        elif ext == ".pdf":
            return {"type": "pdf", "content": content}
        
        elif ext in (".jpg", ".png"):
            return {"type": "image", "content": content}
        
        else:
            st.error(f"Unsupported file type: {ext}")
            return None
    
    except Exception as e:
        st.error(f"Error processing {filename}: {e}")
        return None

def create_pdf_from_content(file_type, content):
    if file_type == "text" or file_type == "docx":
        text_pdf = fitz.open()
        text_page = text_pdf.new_page()
        text_page.insert_textbox(fitz.Rect(50, 50, 550, 800), content, fontsize=12)
        return text_pdf
    
    elif file_type == "xlsx":
        df_excel = pd.read_excel(io.BytesIO(content), engine='openpyxl')
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
        plt.close(fig)
        return fitz.open("pdf", img_bytes.getvalue())
    
    elif file_type == "pdf":
        return fitz.open(stream=content, filetype="pdf")
    
    elif file_type == "image":
        img = Image.open(io.BytesIO(content)).convert("RGB")
        img_bytes = BytesIO()
        img.save(img_bytes, format="PDF")
        return fitz.open("pdf", img_bytes.getvalue())

def main():
    st.title("📄 PDF Merger Tool")

    uploaded_files = st.file_uploader(
        "Upload Files",
        type=["txt", "docx", "pdf", "jpg", "png", "xlsx"],
        accept_multiple_files=True
    )

    if uploaded_files:
        file_contents = []
        for f in uploaded_files:
            processed_file = process_file(f)
            if processed_file:
                file_contents.append({"name": f.name, "processed": processed_file, "order": 0})
        
        st.write("### Reorder Files")
        for i, file_data in enumerate(file_contents):
            file_data["order"] = st.number_input(f"Order for {file_data['name']}", min_value=1, value=i + 1, step=1)
        
        if st.button("Merge Files"):
            try:
                file_contents.sort(key=lambda x: x["order"])
                
                pdf_docs = []
                merged_text = ""

                for file_data in file_contents:
                    processed = file_data["processed"]
                    
                    if processed["type"] == "text":
                        merged_text += processed["content"] + "\n\n"
                    else:
                        pdf_doc = create_pdf_from_content(processed["type"], processed["content"])
                        pdf_docs.append(pdf_doc)

                if merged_text:
                    text_pdf = fitz.open()
                    text_page = text_pdf.new_page()
                    text_page.insert_textbox(fitz.Rect(50, 50, 550, 800), merged_text, fontsize=12)
                    pdf_docs.insert(0, text_pdf)

                if not pdf_docs:
                    st.error("⚠️ No valid files to merge.")
                    return

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

            except Exception as e:
                st.error(f"Unexpected error: {e}")

if __name__ == "__main__":
    main()
