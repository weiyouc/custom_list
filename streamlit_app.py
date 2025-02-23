import streamlit as st
from validator import ExcelValidator
from pathlib import Path
import tempfile
import os
import subprocess

def save_uploaded_file(uploaded_file):
    """Save uploaded file to temporary directory and return path"""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            return tmp_file.name
    except Exception as e:
        st.error(f"Error saving uploaded file: {str(e)}")
        return None

def normalize_files(input_path, shipping_path):
    """Normalize input and shipping files before validation"""
    try:
        # Create normalized file paths
        normalized_input = str(Path(input_path).parent / f"{Path(input_path).stem}_normalized.xlsx")
        normalized_shipping = str(Path(shipping_path).parent / f"{Path(shipping_path).stem}_normalized.xlsx")
        
        # Step 0: Normalize input Excel file
        st.write("Normalizing input file...")
        subprocess.check_call([
            "python", 
            "normalize-inputexcel.py",
            input_path,
            normalized_input
        ])
        
        # Step 1: Normalize shipping list
        st.write("Normalizing shipping file...")
        subprocess.check_call([
            "python",
            "normalize-shipping.py", 
            shipping_path,
            normalized_shipping
        ])
        
        return normalized_input, normalized_shipping
        
    except Exception as e:
        st.error(f"Normalization failed: {str(e)}")
        return None, None

def main():
    st.title("Excel Data Validator")
    st.write("Upload your Excel files to validate the data")
    
    # File uploaders
    input_file = st.file_uploader("Upload Input Excel file (to be validated)", type=['xlsx'])
    shipping_file = st.file_uploader("Upload Shipping List file", type=['xlsx'])
    duty_file = st.file_uploader("Upload Duty Rates file", type=['xlsx'])
    
    if input_file and shipping_file and duty_file:
        if st.button("Validate Files"):
            try:
                with st.spinner("Processing files..."):
                    # Save uploaded files
                    input_path = save_uploaded_file(input_file)
                    shipping_path = save_uploaded_file(shipping_file)
                    duty_path = save_uploaded_file(duty_file)
                    
                    if not all([input_path, shipping_path, duty_path]):
                        st.error("Error saving uploaded files")
                        return
                    
                    # Normalize files first
                    normalized_input, normalized_shipping = normalize_files(input_path, shipping_path)
                    
                    if not normalized_input or not normalized_shipping:
                        st.error("File normalization failed")
                        return
                    
                    # Create validator instance with normalized files
                    validator = ExcelValidator(
                        input_file=normalized_input,
                        shipping_list=normalized_shipping,
                        duty_file=duty_path
                    )
                    
                    # Run validation
                    validator.validate_all()
                    
                    # Generate report
                    validator.generate_report()
                    
                    # Get report path (same directory as input file)
                    report_path = Path(normalized_input).parent / 'validation_report.xlsx'
                    
                    # Offer download if report exists
                    if report_path.exists():
                        with open(report_path, "rb") as file:
                            st.download_button(
                                label="Download Validation Report",
                                data=file,
                                file_name="validation_report.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    else:
                        st.error("Validation report was not generated")
                
                # Cleanup temporary files
                for path in [input_path, shipping_path, duty_path, 
                           normalized_input, normalized_shipping]:
                    if path and Path(path).exists():
                        try:
                            os.unlink(path)
                        except Exception as e:
                            print(f"Error removing temporary file {path}: {e}")
                
                if report_path.exists():
                    try:
                        os.unlink(report_path)
                    except Exception as e:
                        print(f"Error removing report file: {e}")
                        
            except Exception as e:
                st.error(f"Error during validation: {str(e)}")
                st.exception(e)  # This will show the full traceback

if __name__ == "__main__":
    main() 