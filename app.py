import streamlit as st
import sys
import platform
from io import BytesIO

st.title("Diagnostic DOCX Converter")
st.write(f"Python Version: {sys.version}")
st.write(f"Platform: {platform.platform()}")

uploaded_file = st.file_uploader("Upload DOCX", type=['docx'])

if uploaded_file:
    st.write(f"File Name: {uploaded_file.name}")
    st.write(f"File Size: {uploaded_file.size} bytes")
    
    try:
        file_contents = BytesIO(uploaded_file.getvalue())
        st.success(f"Successfully read {file_contents.getbuffer().nbytes} bytes")
    except Exception as e:
        st.error(f"Error reading file: {e}")
