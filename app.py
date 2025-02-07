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
    file_data = []
    for i, f in enumerate(uploaded_files):
        file_data.append({"Order": i + 1, "Filename": f.name}) # Store only order and filename
    df = pd.DataFrame(file_data)

    df['Order'] = df['Order'].astype(int)

    st.write("### Reorder Files")
    st.write("Edit the 'Order' column (lower numbers will appear first).")

    try:
        edited_df = st.data_editor(df, use_container_width=True)
    except pa.lib.ArrowInvalid as e:
        st.error(f"ArrowInvalid Error: {e}")
        st.write("Check if the 'Order' column contains only valid integers.")
        st.stop()

    sorted_df = edited_df.sort_values("Order")
    ordered_filenames = sorted_df["Filename"].tolist()  # Get the ordered filenames

    if st.button("Merge Files"):
        pdf_docs = []
        merged_text = ""

        for filename in ordered_filenames: # Iterate through filenames
            file_obj = next((f for f in uploaded_files if f.name == filename), None) # Find the file object
            if not file_obj:
                continue # Skip if file object not found (shouldn't happen)

            ext = os.path.splitext(filename)[1].lower()

            # ... (Rest of the file processing code from my previous correct response - remains the same)
