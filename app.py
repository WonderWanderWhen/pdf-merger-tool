import streamlit as st
import sys
import platform

st.title("Diagnostic DOCX Converter")

st.write(f"Python Version: {sys.version}")
st.write(f"Platform: {platform.platform()}")

try:
    import docx
    st.write("python-docx: Installed successfully")
except ImportError as e:
    st.error(f"python-docx import failed: {e}")

try:
    import streamlit
    st.write(f"Streamlit Version: {streamlit.__version__}")
except ImportError as e:
    st.error(f"Streamlit import failed: {e}")

uploaded_file = st.file_uploader("Upload DOCX", type=['docx'])

if uploaded_file:
    st.write(f"File Name: {uploaded_file.name}")
    st.write(f"File Size: {uploaded_file.size} bytes")
    
    try:
        file_contents = uploaded_file.getvalue()
        st.success(f"Successfully read {len(file_contents)} bytes")
    except Exception as e:
        st.error(f"Error reading file: {e}")
