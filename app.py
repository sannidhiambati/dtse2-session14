# -*- coding: utf-8 -*-
"""
Streamlit Web App for Universal File-to-Text Conversion.

This script creates a simple web interface where users can:
1. Upload a file (.docx, .xlsx, .pptx, .pdf, .html) via a drag-and-drop interface.
2. The app processes the file into plain text or Markdown.
3. It displays a preview of the first 1000 characters of the extracted text.
4. It provides a download button to save the full text as a markdown file.
5. It shows a rendered preview of the markdown output.
"""

# STEP 1: Import necessary libraries
import streamlit as st
import os
import tempfile
from docx import Document
from openpyxl import load_workbook
from pptx import Presentation
from PyPDF2 import PdfReader
from markdownify import markdownify as md

def convert_file_to_text(filepath):
    """
    Converts the content of a given file to plain text or Markdown.
    This function is adapted from the original Colab script.

    Args:
        filepath (str): The path to the file to be converted.

    Returns:
        str: The extracted text content, or an error string if conversion fails.
    """
    try:
        _, extension = os.path.splitext(filepath)
        extension = extension.lower()
        text_content = ""

        if extension == '.docx':
            document = Document(filepath)
            for paragraph in document.paragraphs:
                text_content += paragraph.text + '\n'
        elif extension == '.xlsx':
            workbook = load_workbook(filepath)
            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                text_content += f"--- Sheet: {sheet_name} ---\n"
                for row in sheet.iter_rows():
                    row_text = [str(cell.value) if cell.value is not None else "" for cell in row]
                    text_content += "\t".join(row_text) + '\n'
        elif extension == '.pptx':
            presentation = Presentation(filepath)
            for slide in presentation.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        text_content += shape.text + '\n'
        elif extension in ['.html', '.htm']:
            with open(filepath, 'r', encoding='utf-8') as f:
                html_content = f.read()
            text_content = md(html_content)
        elif extension == '.pdf':
            reader = PdfReader(filepath)
            for page in reader.pages:
                text_content += page.extract_text() + '\n'
        else:
            # This case handles unsupported file types gracefully.
            return f"Error: Unsupported file type '{extension}'. Please upload a supported file."

        return text_content

    except Exception as e:
        # Return a user-friendly error message.
        return f"Error processing file: {e}. Please ensure the file is not corrupted."

# --- Streamlit App UI ---

# Set the title for the Streamlit app
st.title("Docs to Markdown Converter")
st.markdown("Drag, drop, and download. It's that simple.")

# [1] Create the file uploader widget
# Modified to accept only the file types our conversion function supports.
uploaded_file = st.file_uploader(
    "Choose a file",
    type=['pdf', 'pptx', 'docx', 'xlsx', 'html', 'htm']
)

# Main application logic
if uploaded_file is not None:
    # Handle large file sizes with a warning.
    # 50MB is used as the threshold here.
    if uploaded_file.size > 50 * 1024 * 1024:
        st.warning("Warning: You have uploaded a large file. Processing may take some time.")

    # Use a temporary directory to safely handle the uploaded file
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_filepath = os.path.join(temp_dir, uploaded_file.name)
        with open(temp_filepath, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Show a spinner while the file is being processed.
        with st.spinner("Converting..."):
            # [2] Process the file into a text string
            full_text = convert_file_to_text(temp_filepath)

    # Check if the conversion was successful before proceeding
    if not full_text.startswith("Error"):
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
        # If an error occurred, show the clear error message to the user.
        st.error(full_text)

