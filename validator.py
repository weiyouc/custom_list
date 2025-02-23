import pandas as pd
import numpy as np
from pathlib import Path
from typing import List, Dict, Tuple
import argparse
import logging
from difflib import SequenceMatcher
import warnings
import difflib  # Add this at the top with other imports
import re
import math
import subprocess
import os

def clean_column_name(name: str) -> str:
    """Handle CR characters and normalize names"""
    return str(name).replace('\r', '').replace('\n', '').strip()

def normalize_sheet_name(name: str) -> str:
    """Clean sheet names for comparison"""
    return name.strip().lower().replace(' ', '').replace('-', '').replace('_', '')

class ExcelValidator:
    def __init__(self, input_file: str, shipping_list: str, duty_file: str):
        self.input_file = input_file
        self.shipping_list = shipping_list
        self.duty_file = duty_file
        self.validation_errors = []
        # Set up logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger(__name__)
        # Suppress openpyxl warnings
        warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
        
    def normalize_sheet_name(self, name: str) -> str:
        """Standardize sheet names for matching"""
        return re.sub(r'[^a-zA-Z0-9]', '', name).lower()

    def is_header_row(self, row: pd.Series) -> bool:
        """Header detection for single-row headers with common variations"""
        # Convert row to clean string values
        clean_row = ' '.join(str(cell).strip() for cell in row.values).lower()
        
        # Required columns with common variations
        required = {
            'item': r'\bitem\b.*\bno',
            'model': r'\bmodel\b.*\bno',
            'part': r'\b(p/n|part\s*no)\b',
            'quantity': r'\bquantity\b.*\b(pcs|pc)\b'
        }
        
        # Check for at least 3 matches using regex
        match_count = sum(1 for pattern in required.values() 
                         if re.search(pattern, clean_row))
        
        return match_count >= 3

    def clean_pn_value(self, pn_value) -> str:
        """Clean individual P/N values"""
        raw_pn = str(pn_value).upper().strip()
        clean_pn = re.sub(r'[^A-Z0-9]', '', raw_pn)
        return clean_pn if clean_pn else 'N/A'

    def create_composite_key(self, row: pd.Series) -> str:
        """Get and clean P/N from a dataframe row (Series)"""
        return self.clean_pn_value(row['P/N'] if 'P/N' in row else 'N/A')

    def extract_valid_data(self, df: pd.DataFrame, file_type: str) -> pd.DataFrame:
        """Extract valid data rows from a DataFrame, skipping non-table content"""
        # Add data snapshot logging
        self.logger.debug(f"Raw data sample for {file_type}:\n{df.head(3).to_string()}")
        
        # Only check first 5 rows for headers
        for idx, row in df.head(5).iterrows():
            if self.is_header_row(row):
                header_row = idx
                break
        else:
            self.logger.error(f"Header not found in {file_type}. First rows:\n{df.head(3).to_string()}")
            return pd.DataFrame()
        
        # Remove all rows before the header
        data_df = df.iloc[header_row:].copy()
        
        # Clean column names (remove whitespace and newlines)
        data_df.columns = [clean_column_name(str(col)) for col in data_df.columns]
        
        # Updated column mappings based on normalization specs
        column_mappings = {
            'P/N': ['P/N', 'Part Number', 'Part No', 'PartNo', '料号'],
            'Item Nos': ['Item Nos.', 'Item Number', '项目编号', 'Item No'],
            'Model Nos': ['Model Nos.', 'Model Number', '型号', 'Model'],
            'Description': ['Description', '产品描述', 'Desc'],
            'Quantity PCS': ['Quantity PCS', 'QTY', '数量', 'Quantity'],
            'Unit Price USD': ['Unit Price USD', 'Price', '单价', 'Unit Price'],
            'Amount USD': ['Amount USD', 'Total Amount', '总金额', 'Amount'],
            'India HS code': ['India HS code', 'HS Code', 'HSN Code'],
            'Duty': ['Duty', 'Duty Rate', '税率'],
            'Welfare': ['Welfare', 'Welfare Tax', '福利税'],
            'IGST': ['IGST', 'GST', '综合税']
        }
        
        # Debug log before standardization
        self.logger.debug(f"\nColumns before standardization: {data_df.columns.tolist()}")
        
        # Standardize column names
        for standard_name, possible_names in column_mappings.items():
            found_col = next((col for col in possible_names if col in data_df.columns), None)
            if found_col:
                data_df = data_df.rename(columns={found_col: standard_name})
        
        # Debug log after standardization
        self.logger.debug(f"\nColumns after standardization: {data_df.columns.tolist()}")
        
        # Add P/N column validation
        if 'P/N' not in data_df.columns:
            self.logger.error(f"Missing P/N column in {file_type}")
            return pd.DataFrame()
        
        # Remove rows where P/N is empty or NaN
        if 'P/N' in data_df.columns:
            data_df = data_df[data_df['P/N'].notna()]
            # Clean P/N values
            data_df['P/N'] = data_df['P/N'].astype(str).str.strip()
        
        # Convert Quantity to numeric, handling any non-numeric values
        if 'Quantity' in data_df.columns:
            data_df['Quantity'] = pd.to_numeric(data_df['Quantity'], errors='coerce')
            data_df = data_df[data_df['Quantity'].notna()]  # Remove rows with invalid quantities
        
        # Add P/N cleaning debug
        if 'P/N' in data_df.columns:
            data_df['Cleaned_P/N'] = data_df['P/N'].apply(self.clean_pn_value)
            self.logger.debug(f"Cleaned P/N samples ({file_type}):\n{data_df[['P/N', 'Cleaned_P/N']].head(10)}")
        
        # Debug log the final processed DataFrame
        self.logger.debug(f"\nProcessed DataFrame for {file_type}:\n{data_df.head()}")
        self.logger.debug(f"Extracted {len(data_df)} valid rows from {file_type}")
        self.logger.debug(f"Final columns: {data_df.columns.tolist()}")
        
        # Remove empty columns (common in files with formatting)
        data_df = data_df.dropna(axis=1, how='all')
        
        if data_df.empty:
            self.logger.warning(f"No data rows found after processing in {file_type}")
            return pd.DataFrame()
        
        # Additional check for meaningful data
        if data_df.iloc[:, 0].isna().all():
            self.logger.warning(f"First column is empty in {file_type}, possible formatting issues")
            return pd.DataFrame()
        
        return data_df

    def process_duty_file(self) -> pd.DataFrame:
        """Process duty file with special handling"""
        duty_df = pd.read_excel(self.duty_file, header=None)
        return self.extract_valid_data(duty_df, "duty file")

    def process_input_file(self) -> Dict[str, pd.DataFrame]:
        """Process input file sheets"""
        input_sheets = {}
        xl = pd.ExcelFile(self.input_file)
        for sheet_name in xl.sheet_names:
            df = pd.read_excel(self.input_file, sheet_name=sheet_name, header=None)
            processed_df = self.extract_valid_data(df, f"input sheet {sheet_name}")
            if not processed_df.empty:
                input_sheets[sheet_name] = processed_df
        return input_sheets

    def load_shipping_data(self, file_path: str) -> Dict[str, pd.DataFrame]:
        """Load shipping data with flexible header detection"""
        all_sheets = pd.read_excel(file_path, sheet_name=None, header=None)
        valid_sheets = {}
        
        for sheet_name, df in all_sheets.items():
            header_row = None
            # Check first 5 rows for headers
            for idx in range(min(5, len(df))):
                if self.is_header_row(df.iloc[idx]):
                    header_row = idx
                    break
            
            if header_row is not None:
                try:
                    # Use found header row
                    df.columns = df.iloc[header_row]
                    valid_df = df.iloc[header_row+1:].dropna(how='all')
                    
                    # Validate we found actual data rows
                    if len(valid_df) > 0 and 'Item No.' in valid_df.columns:
                        valid_sheets[sheet_name] = valid_df.reset_index(drop=True)
                except Exception as e:
                    self.logger.error(f"Error processing {sheet_name}: {str(e)}")
        
        return valid_sheets

    def load_duty_rates(self, file_path: str) -> pd.DataFrame:
        """Load duty rate file with proper header detection"""
        df = pd.read_excel(file_path, header=None)
        
        # Find header row
        header_row = None
        for idx, row in df.iterrows():
            if 'Item name' in row.values and 'India HS code' in row.values:
                header_row = idx
                break
        
        if header_row is not None:
            df.columns = df.iloc[header_row]
            return df.iloc[header_row+1:].dropna(how='all')
        else:
            self.logger.error("Could not find header in duty rate file")
            return pd.DataFrame()

    def load_excel_files(self) -> dict:
        """Load all Excel files with proper multi-sheet handling"""
        try:
            # Load duty rates first
            self.duty_rates = self.load_duty_rates(self.duty_file)
            
            # Keep original sheet names
            input_sheets = pd.read_excel(self.input_file, sheet_name=None)
            return {
                'shipping': self.load_shipping_data(self.shipping_list),
                'input_data': {self.normalize_sheet_name(name): df 
                              for name, df in input_sheets.items()},
                'input_sheets': input_sheets,
                'duty_rates': self.duty_rates  # Now properly initialized
            }
        except Exception as e:
            self.logger.error(f"Error loading files: {str(e)}")
            raise

    def validate_sheet(self, sheet_df: pd.DataFrame, sheet_name: str,
                      shipping_df: pd.DataFrame, duty_df: pd.DataFrame):
        # Build shipping index using cleaned P/N values
        shipping_df['clean_pn'] = shipping_df['P/N'].apply(self.clean_pn_value)
        shipping_index = shipping_df.set_index('clean_pn').to_dict('index')
        
        for idx, input_row in sheet_df.iterrows():
            # Get and clean P/N from input
            input_pn = self.clean_pn_value(input_row.get('P/N', 'N/A'))
            
            if input_pn == 'N/A':
                self.log_error(sheet_name, idx, input_pn, "Missing P/N in input row")
                continue
            
            # Find in shipping list using cleaned P/N
            shipping_row = shipping_index.get(input_pn)
            
            if not shipping_row:
                # Try matching with different P/N formats
                alt_pn = self.clean_pn_value(input_pn.replace('.', ''))  # Try without dots
                shipping_row = shipping_index.get(alt_pn)
                
                if not shipping_row:
                    self.log_error(sheet_name, idx, input_pn, "No matching shipping entry for P/N")
                    continue
            
            # Step 2.3: Validate columns
            self.validate_columns(input_row, shipping_row, [
                'Item Nos', 'Model Nos', 'Description',
                'Quantity PCS', 'Unit Price USD', 'Amount USD'
            ], sheet_name, idx)
            
            # Step 3: Validate duty info
            self.validate_duty_info(input_row, sheet_name, idx)

    def validate_columns(self, input_row: pd.Series, shipping_row: dict, columns: list, 
                        sheet_name: str, row_idx: int):
        """Validate column values between input and shipping data"""
        for col in columns:
            input_val = input_row.get(col, 'N/A')
            shipping_val = shipping_row.get(col, 'N/A')
            
            # Handle numeric comparisons
            if isinstance(input_val, (int, float)) and isinstance(shipping_val, (int, float)):
                if not math.isclose(input_val, shipping_val, rel_tol=0.01):
                    self.log_error(
                        sheet_name,
                        row_idx,
                        self.clean_pn_value(input_row.get('P/N', 'N/A')),
                        f"Value mismatch in column {col}: {input_val} vs {shipping_val}"
                    )
            else:
                # Handle text comparisons
                self.validate_text(input_row, shipping_row, col, sheet_name, row_idx)

    def validate_duty_info(self, input_row: pd.Series, sheet_name: str, row_idx: int):
        """Validate duty rate information against loaded rates"""
        if self.duty_rates.empty:
            self.logger.warning("Skipping duty validation - no duty data loaded")
            return
        
        # Get item name from input
        item_name = str(input_row.get('Item name', 'N/A')).strip()
        if item_name == 'N/A':
            self.log_error(sheet_name, row_idx, item_name, "Missing item name")
            return
        
        # Escape special characters in item_name for regex
        pattern = re.escape(item_name)
        match = self.duty_rates[self.duty_rates['Item name'].str.contains(
            pattern, case=False, na=False, regex=True)]
        
        if not match.empty:
            # Handle numeric HS codes and formatting variations
            hs_code = input_row.get('India HS code', 'N/A')
            
            # Convert to string and clean
            if isinstance(hs_code, (int, float)):
                # Handle numeric values: 1234.0 -> "1234", 1234.5 -> "1234.5"
                hs_str = f"{hs_code:.10f}".rstrip('0').rstrip('.') if '.' in str(hs_code) else str(int(hs_code))
            else:
                hs_str = str(hs_code).strip()
            
            # More flexible regex pattern
            if not re.match(r'^(\d{4,10}(\.\d{1,10})?|\d+-\d+)$', hs_str):
                self.log_error(sheet_name, row_idx, hs_str, 
                              f"Invalid HS Code format: {hs_str} (accepts numbers, decimals, or hyphenated formats)")
        else:
            self.log_error(sheet_name, row_idx, item_name, "No matching duty rate found")

    def validate_text(self, input_row, reference_row, col_name, 
                     sheet_name: str, row_idx: int, threshold=0.85):
        # Ensure we're comparing strings by explicitly converting
        input_text = str(input_row.get(col_name, '')).lower()
        ref_text = str(reference_row.get(col_name, '')).lower()
        
        similarity = SequenceMatcher(None, input_text, ref_text).ratio()
        if similarity < threshold:
            self.log_error(
                sheet_name,
                row_idx,
                self.clean_pn_value(input_row.get('P/N', 'N/A')),
                f"Text similarity low in column {col_name}: {input_text} vs {ref_text}"
            )

    def validate_all(self):
        data = self.load_excel_files()
        
        # Store original sheet names for reporting
        matched_pairs = []
        for input_name, input_df in data['input_data'].items():
            original_name = self.get_original_sheet_name(input_name, data['input_sheets'])
            best_match = None
            best_score = 0
            
            for shipping_name in data['shipping'].keys():
                score = SequenceMatcher(
                    None, 
                    self.normalize_sheet_name(input_name),
                    self.normalize_sheet_name(shipping_name)
                ).ratio()
                
                if score > best_score and score > 0.6:
                    best_score = score
                    best_match = shipping_name
                    
            if best_match:
                matched_pairs.append((input_df, original_name, best_match))

        # Validate matched pairs
        for input_df, original_sheet_name, shipping_name in matched_pairs:
            try:
                shipping_df = data['shipping'][shipping_name]
                # Pass original sheet name to validation
                self.validate_sheet(input_df, original_sheet_name, shipping_df, data['duty_rates'])
            except Exception as e:
                self.logger.error(f"Validation failed for {original_sheet_name}: {str(e)}")

    def generate_report(self):
        """
        Generate Excel report with validation errors
        """
        # Convert validation errors to DataFrame
        error_df = pd.DataFrame(self.validation_errors)
        
        # Save to Excel file
        output_path = Path(self.input_file).parent / 'validation_report.xlsx'
        error_df.to_excel(output_path, index=False)
        print(f"Validation report generated: {output_path}")

    def log_error(self, sheet_name: str, row_idx: int, pn: str, error_msg: str):
        """Log validation errors with proper row numbers"""
        self.validation_errors.append({
            'Sheet': sheet_name,
            'Row': row_idx + 1,  # Convert 0-based to 1-based
            'P/N': pn,
            'Error': error_msg
        })
        self.logger.debug(f"Validation error in {sheet_name} row {row_idx+1}: {error_msg}")

    def get_original_sheet_name(self, normalized_name: str, original_sheets: dict) -> str:
        """Find original sheet name from normalized version"""
        for name in original_sheets.keys():
            if self.normalize_sheet_name(name) == normalized_name:
                return name
        return normalized_name  # Fallback if not found

def main():
    # Set up argument parser
    parser = argparse.ArgumentParser(
        description='Validate Excel files against shipping list and duty rates',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    python excel_validator.py input.xlsx shipping_list.xlsx duty_rates.xlsx
    python excel_validator.py input.xlsx shipping_list.xlsx duty_rates.xlsx --debug

Note: The validation report will be generated as 'validation_report.xlsx' in the same directory as the input file.
        """
    )
    
    parser.add_argument('input_file', type=str, 
                       help='Path to the input Excel file to be validated')
    parser.add_argument('shipping_list', type=str, 
                       help='Path to the shipping list Excel file containing reference data')
    parser.add_argument('duty_file', type=str, 
                       help='Path to the duty rates Excel file containing tax information')
    parser.add_argument('--debug', action='store_true', 
                       help='Enable debug logging for detailed execution information')

    args = parser.parse_args()
    
    # Set debug level if requested
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    
    # Add normalization step before validation
    try:
        # Create output paths with _normalized suffix
        input_path = Path(args.input_file)
        normalized_input_path = input_path.with_stem(f"{input_path.stem}_normalized")
        
        shipping_path = Path(args.shipping_list)
        normalized_shipping_path = shipping_path.with_stem(f"{shipping_path.stem}_normalized")

        # Step 0: Normalize input Excel file
        print(f"Normalizing input file to {normalized_input_path}...")
        subprocess.check_call([
            "python", 
            "normalize-inputexcel.py",
            str(input_path),
            str(normalized_input_path)
        ])
        
        # Verify normalization output
        if not normalized_input_path.exists():
            raise FileNotFoundError(f"Normalization failed to create {normalized_input_path}")

        # Step 1: Normalize shipping list
        print(f"Normalizing shipping file to {normalized_shipping_path}...")
        subprocess.check_call([
            "python",
            "normalize-shipping.py", 
            str(shipping_path),
            str(normalized_shipping_path)
        ])

        if not normalized_shipping_path.exists():
            raise FileNotFoundError(f"Normalization failed to create {normalized_shipping_path}")

    except Exception as e:
        logging.error(f"Normalization failed: {str(e)}")
        return

    # Create and run validator with normalized files
    validator = ExcelValidator(
        input_file=normalized_input_path,
        shipping_list=normalized_shipping_path,
        duty_file=args.duty_file
    )
    validator.validate_all()
    validator.generate_report()

if __name__ == "__main__":
    main() 