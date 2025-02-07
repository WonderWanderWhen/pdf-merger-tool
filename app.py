import streamlit as st
import io
import os
import docx
import fitz
import pandas as pd
from PIL import Image
import matplotlib.pyplot as plt

def convert_to_pdf(file, filename):
    ext = os.path.splitext(filename)[1].lower()
    
    try:
        if ext == '.txt':
            text = file.decode('utf-8')
            pdf = fitz.open()
            page = pdf.new_page()
            page.insert_textbox(fitz.Rect(50, 50, 550, 800), text, fontsize=12)
            return pdf
        
        elif ext == '.docx':
            doc = docx.Document(io.BytesIO(file))
            text = '\n'.join([para.text for para in doc.paragraphs])
            pdf = fitz.open()
            page = pdf.new_page()
            page.insert_textbox(fitz.Rect(50, 50, 550, 800), text, fontsize=12)
            return pdf
        
        elif ext == '.xlsx':
            df = pd.read_excel(io.BytesIO(file), engine='openpyxl')
            fig, ax = plt.subplots(figsize=(8, 6))
            ax.table(cellText=df.values, colLabels=df.columns, loc='center')
            ax.axis('off')
            plt.tight_layout()
            
            img_buf = io.BytesIO()
            plt.savefig(img_buf, format='png', dpi=300)
            img_buf.seek(0)
            img = Image.open(img_buf).convert('RGB')
            
            img_bytes = io.BytesIO()
            img.save(img_bytes, format='PDF')
            plt.close(fig)
            
            return fitz.open('pdf', img_bytes.getvalue())
        
        elif ext == '.pdf':
            return fitz.open(stream=file, filetype='pdf')
        
        elif ext in ['.jpg', '.png']:
            img = Image.open(io.BytesIO(file)).convert('RGB')
            img_bytes = io.BytesIO()
            img.save(img_bytes, format='PDF')
            return fitz.open('pdf', img_bytes.getvalue())
        
    except Exception as e:
        st.error(f"Error converting {filename}: {e}")
        return None

def main():
    st.title("📄 PDF Merger Tool")

    uploaded_files = st.file_uploader(
        "Upload Files", 
        type=["txt", "docx", "pdf", "jpg", "png", "xlsx"], 
        accept_multiple_files=True
    )

    if uploaded_files:
        file_details = []
        for f in uploaded_files:
            file_details.append({
                'name': f.name, 
                'content': f.getvalue(), 
                'order': 0
            })

        st.write("### Reorder Files")
        for i, file_data in enumerate(file_details):
            file_data['order'] = st.number_input(f"Order for {file_data['name']}", min_value=1, value=i+1, step=1)

        if st.button("Merge Files"):
            file_details.sort(key=lambda x: x['order'])
            
            pdf_docs = []
            for file_data in file_details:
                pdf = convert_to_pdf(file_data['content'], file_data['name'])
                if pdf:
                    pdf_docs.append(pdf)

            if pdf_docs:
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
            else:
                st.error("No valid files to merge")

if __name__ == "__main__":
    main()
