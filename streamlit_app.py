import streamlit as st
from validator import ExcelValidator
from pathlib import Path
import tempfile
import os
import subprocess

# Define translations
TRANSLATIONS = {
    'English': {
        'title': "Excel Data Validator",
        'description': "Upload your Excel files to validate the data",
        'language_selector': "Language",
        'input_file_label': "Upload Input Excel file (to be validated)",
        'shipping_file_label': "Upload Shipping List file",
        'duty_file_label': "Upload Duty Rates file",
        'validate_button': "Validate Files",
        'processing': "Processing files...",
        'normalizing_input': "Normalizing input file...",
        'normalizing_shipping': "Normalizing shipping file...",
        'download_report': "Download Validation Report",
        'error_saving': "Error saving uploaded files",
        'error_normalization': "File normalization failed",
        'error_no_report': "Validation report was not generated",
        'error_validation': "Error during validation: {}"
    },
    '中文': {
        'title': "Excel数据验证器",
        'description': "上传Excel文件以验证数据",
        'language_selector': "语言",
        'input_file_label': "上传输入Excel文件（待验证）",
        'shipping_file_label': "上传装运清单文件",
        'duty_file_label': "上传税率文件",
        'validate_button': "验证文件",
        'processing': "正在处理文件...",
        'normalizing_input': "正在标准化输入文件...",
        'normalizing_shipping': "正在标准化装运文件...",
        'download_report': "下载验证报告",
        'error_saving': "保存上传文件时出错",
        'error_normalization': "文件标准化失败",
        'error_no_report': "未生成验证报告",
        'error_validation': "验证过程中出错: {}"
    }
}

def get_text(key):
    """Get translated text based on selected language"""
    lang = st.session_state.get('language', 'English')
    return TRANSLATIONS[lang][key]

def save_uploaded_file(uploaded_file):
    """Save uploaded file to temporary directory and return path"""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            return tmp_file.name
    except Exception as e:
        st.error(f"{get_text('error_saving')}: {str(e)}")
        return None

def normalize_files(input_path, shipping_path):
    """Normalize input and shipping files before validation"""
    try:
        # Create normalized file paths
        normalized_input = str(Path(input_path).parent / f"{Path(input_path).stem}_normalized.xlsx")
        normalized_shipping = str(Path(shipping_path).parent / f"{Path(shipping_path).stem}_normalized.xlsx")
        
        # Step 0: Normalize input Excel file
        st.write(get_text('normalizing_input'))
        subprocess.check_call([
            "python", 
            "normalize-inputexcel.py",
            input_path,
            normalized_input
        ])
        
        # Step 1: Normalize shipping list
        st.write(get_text('normalizing_shipping'))
        subprocess.check_call([
            "python",
            "normalize-shipping.py", 
            shipping_path,
            normalized_shipping
        ])
        
        return normalized_input, normalized_shipping
        
    except Exception as e:
        st.error(f"{get_text('error_normalization')}: {str(e)}")
        return None, None

def main():
    # Language selector in sidebar
    if 'language' not in st.session_state:
        st.session_state.language = 'English'
    
    # Create a container for the language selector
    with st.container():
        col1, col2 = st.columns([6, 1])
        with col2:
            selected_lang = st.selectbox(
                get_text('language_selector'),
                options=['English', '中文'],
                key='language'
            )
    
    st.title(get_text('title'))
    st.write(get_text('description'))
    
    # File uploaders
    input_file = st.file_uploader(get_text('input_file_label'), type=['xlsx'])
    shipping_file = st.file_uploader(get_text('shipping_file_label'), type=['xlsx'])
    duty_file = st.file_uploader(get_text('duty_file_label'), type=['xlsx'])
    
    if input_file and shipping_file and duty_file:
        if st.button(get_text('validate_button')):
            try:
                with st.spinner(get_text('processing')):
                    # Save uploaded files
                    input_path = save_uploaded_file(input_file)
                    shipping_path = save_uploaded_file(shipping_file)
                    duty_path = save_uploaded_file(duty_file)
                    
                    if not all([input_path, shipping_path, duty_path]):
                        st.error(get_text('error_saving'))
                        return
                    
                    # Normalize files first
                    normalized_input, normalized_shipping = normalize_files(input_path, shipping_path)
                    
                    if not normalized_input or not normalized_shipping:
                        st.error(get_text('error_normalization'))
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
                    
                    # Get report path
                    report_path = Path(normalized_input).parent / 'validation_report.xlsx'
                    
                    # Offer download if report exists
                    if report_path.exists():
                        with open(report_path, "rb") as file:
                            st.download_button(
                                label=get_text('download_report'),
                                data=file,
                                file_name="validation_report.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    else:
                        st.error(get_text('error_no_report'))
                
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
                st.error(get_text('error_validation').format(str(e)))
                st.exception(e)

if __name__ == "__main__":
    main() 