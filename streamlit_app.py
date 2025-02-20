import streamlit as st
from excel_validator import ExcelValidator
from pathlib import Path
import tempfile
import os

def save_uploaded_file(uploaded_file):
    """Save uploaded file to temporary directory and return path"""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            return Path(tmp_file.name)
    except Exception as e:
        st.error(f"Error saving uploaded file: {str(e)}")
        return None

def main():
    st.title("Excel Data Validator")
    st.write("Upload your Excel files to validate the data")
    
    # File uploaders
    check_file = st.file_uploader("Upload Check List file", type=['xlsx'])
    bom_file = st.file_uploader("Upload BOM file", type=['xlsx'])
    tax_file = st.file_uploader("Upload Tax file", type=['xlsx'])
    
    if check_file and bom_file and tax_file:
        if st.button("Validate Files"):
            try:
                # Save uploaded files
                check_path = save_uploaded_file(check_file)
                bom_path = save_uploaded_file(bom_file)
                tax_path = save_uploaded_file(tax_file)
                
                # Create temporary output file
                output_path = Path(tempfile.mktemp(suffix='.xlsx'))
                
                # Initialize validator
                validator = ExcelValidator()
                
                with st.spinner("Validating data..."):
                    # Load and validate data
                    df_check, df_bom, df_tax = validator.load_files(check_path, bom_path, tax_path)
                    validator.validate_data(df_check, df_bom, df_tax)
                    
                    # Generate report
                    validator.generate_report(output_path)
                    
                    # Offer download of report
                    with open(output_path, "rb") as file:
                        st.download_button(
                            label="Download Validation Report",
                            data=file,
                            file_name="validation_report.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                
                # Cleanup temporary files
                for path in [check_path, bom_path, tax_path, output_path]:
                    try:
                        os.unlink(path)
                    except OSError as e:
                        print(f"Error removing temporary file {path}: {e}")
                        
            except Exception as e:
                st.error(f"Error during validation: {str(e)}")

if __name__ == "__main__":
    main() 