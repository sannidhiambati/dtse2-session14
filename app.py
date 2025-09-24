# -*- coding: utf-8 -*-
"""
Streamlit Web App for Universal File-to-Text Conversion.

This script creates a simple web interface where users can:
1. Upload a file (e.g., .docx, .xlsx, .pptx, .pdf, images, .mp3) via a drag-and-drop interface.
2. The app processes the file into plain text (or Markdown for HTML).
3. It displays a preview of the first 1000 characters of the extracted text.
4. It provides a download button for the user to save the full text as a markdown file.
5. It shows a rendered preview of the markdown output.
"""

# STEP 1: Import necessary libraries
import streamlit as st
import os
import textract
from markdownify import markdownify as md
import tempfile # To handle temporary file storage

def convert_file_to_text(filepath):
    """
    Converts the content of a given file to plain text or Markdown.
    This function is adapted from the original Colab script.

    Args:
        filepath (str): The path to the file to be converted.

    Returns:
        str: The extracted text content, or an empty string if conversion fails.
    """
    try:
        # 'textract.process()' handles the conversion of various file types.
        # It returns the content as bytes, so we decode it to a UTF-8 string.
        byte_content = textract.process(filepath)
        text_content = byte_content.decode('utf-8')

        # Check the file extension to decide if it needs HTML-to-Markdown conversion.
        _, extension = os.path.splitext(filepath)
        if extension.lower() in ['.html', '.htm']:
            # Convert the extracted HTML into clean, readable Markdown.
            return md(text_content)
        else:
            # For all other file types, return the plain text.
            return text_content
    except Exception as e:
        # Return an error message if text extraction fails.
        # This will be displayed to the user in the Streamlit interface.
        return f"Error processing file: {e}. Please ensure the file is not corrupted and is a supported format."

# --- Streamlit App UI ---

# Set the title for the Streamlit app
st.title("Docs to Markdown Converter")
st.markdown("Drag, drop, and download. It's that simple.")

# [1] Create the file uploader widget
# This allows users to drag and drop files directly into the browser.
# Modified to accept only the specified file types.
uploaded_file = st.file_uploader(
    "Choose a file",
    type=['pdf', 'pptx', 'docx', 'xlsx', 'jpg', 'jpeg', 'png', 'mp3']
)

# Main application logic
if uploaded_file is not None:
    # Handle large file sizes gracefully with a warning.
    # Check if file size is greater than 50MB (50 * 1024 * 1024 bytes)
    if uploaded_file.size > 50 * 1024 * 1024:
        st.warning("Warning: You have uploaded a large file. Processing may take some time.")

    # Use a temporary directory to safely handle the uploaded file
    with tempfile.TemporaryDirectory() as temp_dir:
        # Create a temporary file path
        temp_filepath = os.path.join(temp_dir, uploaded_file.name)

        # Write the uploaded file's bytes to the temporary file
        with open(temp_filepath, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Show a spinner with the new message while the file is being processed.
        with st.spinner("Converting..."):
            # [2] Process the file into a text string
            full_text = convert_file_to_text(temp_filepath)

    # Check if the conversion was successful before proceeding
    if not full_text.startswith("Error processing file"):
        st.success("File processed successfully!")

        # [3] Offer a preview display of the first 1000 characters
        st.subheader("Preview (First 1000 characters)")
        st.text_area("Output Preview", full_text[:1000], height=250, disabled=True)

        # Prepare a filename for the downloadable markdown file
        output_filename = f"converted_{os.path.splitext(uploaded_file.name)[0]}.md"

        # [4] Offer the output file for download as a markdown file.
        st.download_button(
           label="Download Markdown File",
           data=full_text,
           file_name=output_filename,
           mime="text/markdown"
        )

        # Show the rendered markdown in an expandable section.
        with st.expander("Rendered Preview"):
            st.markdown(full_text, unsafe_allow_html=True)

    else:
        # If an error occurred during conversion, show it to the user
        st.error(full_text)

