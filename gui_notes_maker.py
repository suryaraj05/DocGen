import streamlit as st
import os
from docx2pdf import convert
from code_exp1 import generate_document  # Import your document generation function

def main():
    st.set_page_config(page_title="Document Generator", layout="wide")
    
    st.title("📝 Document Generator")
    
    # User input for document name
    doc_name = st.text_input("Enter the name for your document:", placeholder="My Document")
    
    # User input for text area
    input_text = st.text_area("Enter your text (or copy from your file):", height=300)
    
    # File uploader for existing text files
    uploaded_file = st.file_uploader("Or upload an input text file", type=["txt"])
    
    # Default output path
    output_path = f"{doc_name}.docx" if doc_name else "output.docx"

    # If a file is uploaded, read the content
    if uploaded_file is not None:
        input_text = uploaded_file.read().decode("utf-8")
        st.success("File uploaded successfully!")

    # Create columns for buttons
    col1, col2, col3 = st.columns(3)

    with col1:
        if st.button("Generate Document"):
            if input_text:
                try:
                    generate_document(input_text, output_path)  # Call the document generation function
                    st.success(f"Document '{output_path}' created successfully!")
                except Exception as e:
                    st.error(f"Error: {str(e)}")
            else:
                st.warning("Please enter text or upload a file.")

    with col2:
        if st.button("Open Document"):
            try:
                os.startfile(output_path)  # For Windows
                st.success(f"Opening document '{output_path}' for review!")
            except Exception as e:
                st.error(f"Could not open document: {str(e)}")

    with col3:
        if st.button("Convert to PDF"):
            try:
                convert(output_path)  # Converts the .docx file to .pdf
                st.success(f"Document converted to PDF: '{output_path.replace('.docx', '.pdf')}'")
            except Exception as e:
                st.error(f"Error converting to PDF: {str(e)}")

    # Delete button
    if st.button("Delete Word Document"):
        try:
            os.remove(output_path)
            st.success(f"Deleted Word document: '{output_path}'")
        except Exception as e:
            st.error(f"Error deleting document: {str(e)}")

if __name__ == "__main__":
    main()
