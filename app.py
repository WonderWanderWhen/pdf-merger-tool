import streamlit as st
import docx
import fitz
import pandas as pd
from PIL import Image
import os
from io import BytesIO
import io
import matplotlib.pyplot as plt
import pyarrow as pa

st.title("📄 PDF Merger Tool")

uploaded_files = st.file_uploader(
    "Upload Files",
    type=["txt", "docx", "pdf", "jpg", "png", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    # Create a dictionary to store filenames and their order
    file_order = {}
    for i, f in enumerate(uploaded_files):
        file_order[f.name] = i + 1  # Store the initial order

    # Create a DataFrame *only* when the "Merge Files" button is clicked
    if st.button("Merge Files"):
        file_data = []
        for filename, order in file_order.items():  # Use stored order
            file_data.append({"Order": order, "Filename": filename})
        df = pd.DataFrame(file_data)
        df['Order'] = df['Order'].astype(int)

        try:
            sorted_df = df.sort_values("Order")  # Sort before data_editor
            ordered_filenames = sorted_df["Filename"].tolist()

        except pa.lib.ArrowInvalid as e:
            st.error(f"ArrowInvalid Error: {e}")
            st.write("Check if the 'Order' column contains only valid integers.")
            st.stop()



        pdf_docs = []
        merged_text = ""

        for filename in ordered_filenames:
            file_obj = next((f for f in uploaded_files if f.name == filename), None)
            if not file_obj:
                continue

            ext = os.path.splitext(filename)[1].lower()

            # ... (Rest of the file processing code - remains the same)
