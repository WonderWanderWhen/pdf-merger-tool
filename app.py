import streamlit as st
import fitz
import io
import docx

def main():
    st.title("📄 PDF Merger Tool")

    uploaded_file = st.file_uploader(
        "Upload File", 
        type=["docx"]
    )

    if uploaded_file:
        try:
            # Read the DOCX file
            doc = docx.Document(io.BytesIO(uploaded_file.getvalue()))
            
            # Extract text
            doc_text = "\n".join([para.text for para in doc.paragraphs])
            
            # Create PDF
            pdf = fitz.open()
            page = pdf.new_page()
            page.insert_textbox(fitz.Rect(50, 50, 550, 800), doc_text, fontsize=12)
            
            # Save PDF
            pdf.save("output.pdf")
            
            # Provide download
            with open("output.pdf", "rb") as f:
                st.download_button(
                    "Download PDF",
                    data=f,
                    file_name="converted.pdf",
                    mime="application/pdf"
                )
        
        except Exception as e:
            st.error(f"Error processing file: {e}")

if __name__ == "__main__":
    main()
