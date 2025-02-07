import streamlit as st
import io
import os
from typing import List, Optional

def main():
    st.title("📄 PDF Converter Tool")

    uploaded_file = st.file_uploader(
        "Upload File", 
        type=["docx"],
        accept_multiple_files=False
    )

    if uploaded_file:
        try:
            # Directly read file bytes
            file_bytes = uploaded_file.read()
            
            # Show basic file info
            st.write(f"File Name: {uploaded_file.name}")
            st.write(f"File Size: {len(file_bytes)} bytes")
            
            st.success("File uploaded successfully!")
        
        except Exception as e:
            st.error(f"Error reading file: {e}")

if __name__ == "__main__":
    main()
