import pandas as pd
import pathlib
from typing import Dict, List, Any
import streamlit as st
from pathlib import Path

class ExcelValidator:
    def __init__(self):
        self.results = {
            'pn_mismatch': [],
            'desc_mismatch': [],
            'hsn_mismatch': [],
            'duty_mismatch': []
        }
        
        # Updated column mappings with more variations
        self.column_mappings = {
            'check_file': {
                'P/N': ['P/N', 'PN', 'PART NUMBER', 'PART NO', 'PART NO.', 'P/N ', 'PART_NUMBER'],
                'DESC': ['DESC', 'DESCRIPTION', 'ITEM DESCRIPTION', 'DESC '],
                'HSN': ['HSN', 'HSN CODE', 'HS CODE', 'HSN '],
                'DUTY': ['DUTY', 'BCD', 'BASIC DUTY', 'DUTY ']
            },
            'bom_file': {
                'P/N': ['P/N', 'PN', 'PART NUMBER', 'MODEL NO', 'PART NO', 'PART NO.', 'P/N '],
                'DESCRIPTION': ['DESCRIPTION', 'DESC', 'ITEM DESCRIPTION', 'DESCRIPTION ']
            },
            'tax_file': {
                'HSN': ['INDIA HS CODE', 'HS CODE', 'HSN', 'HSN CODE', 'INDIA HS CODE '],
                'BCD': ['BCD', 'DUTY', 'BASIC DUTY', 'BCD ']
            }
        }
        
    def find_column_name(self, df: pd.DataFrame, possible_names: List[str]) -> str:
        """Find the actual column name from possible names"""
        df_columns = set(df.columns)
        # Debug print
        print(f"Available columns: {df_columns}")
        print(f"Looking for one of these columns: {possible_names}")
        
        # First try exact match
        for name in possible_names:
            if name in df_columns:
                return name
        
        # If no exact match, try case-insensitive match
        df_columns_lower = {col.lower() for col in df_columns}
        for name in possible_names:
            if name.lower() in df_columns_lower:
                # Find the original column name with matching lowercase
                for col in df_columns:
                    if col.lower() == name.lower():
                        return col
        
        raise ValueError(f"Could not find any of these columns: {possible_names}\nAvailable columns are: {sorted(df_columns)}")
        
    def clean_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """Clean DataFrame by converting columns to uppercase and stripping whitespace"""
        # Before cleaning, print original column names
        print("Original column names:", df.columns.tolist())
        
        # Clean column names
        df.columns = df.columns.str.strip().str.upper()
        
        # After cleaning, print cleaned column names
        print("Cleaned column names:", df.columns.tolist())
        
        # Clean data in columns
        for col in df.select_dtypes(include=['object']):
            df[col] = df[col].astype(str).str.strip().str.upper()
        return df
    
    def load_files(self, check_file: Path, bom_file: Path, tax_file: Path) -> tuple:
        """Load and clean Excel files"""
        try:
            df_check = pd.read_excel(check_file, engine='openpyxl')
            df_bom = pd.read_excel(bom_file, engine='openpyxl')
            df_tax = pd.read_excel(tax_file, engine='openpyxl')
            
            # Clean data
            df_check = self.clean_data(df_check)
            df_bom = self.clean_data(df_bom)
            df_tax = self.clean_data(df_tax)
            
            # Print column names for debugging
            print("Check file columns:", df_check.columns.tolist())
            print("BOM file columns:", df_bom.columns.tolist())
            print("Tax file columns:", df_tax.columns.tolist())
            
            return (df_check, df_bom, df_tax)
        except Exception as e:
            raise Exception(f"Error loading files: {str(e)}")
    
    def validate_data(self, df_check: pd.DataFrame, df_bom: pd.DataFrame, df_tax: pd.DataFrame):
        """Perform validation checks"""
        try:
            # Get column names for each file
            check_pn = self.find_column_name(df_check, self.column_mappings['check_file']['P/N'])
            check_desc = self.find_column_name(df_check, self.column_mappings['check_file']['DESC'])
            check_hsn = self.find_column_name(df_check, self.column_mappings['check_file']['HSN'])
            check_duty = self.find_column_name(df_check, self.column_mappings['check_file']['DUTY'])
            
            bom_pn = self.find_column_name(df_bom, self.column_mappings['bom_file']['P/N'])
            bom_desc = self.find_column_name(df_bom, self.column_mappings['bom_file']['DESCRIPTION'])
            
            tax_hsn = self.find_column_name(df_tax, self.column_mappings['tax_file']['HSN'])
            tax_bcd = self.find_column_name(df_tax, self.column_mappings['tax_file']['BCD'])
            
            # Get reference values
            bom_pns = set(df_bom[bom_pn].astype(str).unique())
            bom_descs = set(df_bom[bom_desc].astype(str).unique())
            tax_hsns = set(df_tax[tax_hsn].astype(str).unique())
            tax_bcds = set(df_tax[tax_bcd].astype(str).unique())
            
            total_rows = len(df_check)
            
            for idx, row in df_check.iterrows():
                # Check P/N
                if str(row[check_pn]).upper().strip() not in bom_pns:
                    self.results['pn_mismatch'].append({
                        'Row': idx + 2,
                        'P/N': row[check_pn],
                        'Status': 'Not found in BOM'
                    })
                
                # Check Description
                if str(row[check_desc]).upper().strip() not in bom_descs:
                    self.results['desc_mismatch'].append({
                        'Row': idx + 2,
                        'Description': row[check_desc],
                        'Status': 'Not found in BOM'
                    })
                
                # Check HSN
                if str(row[check_hsn]).upper().strip() not in tax_hsns:
                    self.results['hsn_mismatch'].append({
                        'Row': idx + 2,
                        'HSN': row[check_hsn],
                        'Status': 'Not found in Tax file'
                    })
                
                # Check Duty
                if str(row[check_duty]).upper().strip() not in tax_bcds:
                    self.results['duty_mismatch'].append({
                        'Row': idx + 2,
                        'Duty': row[check_duty],
                        'Status': 'Not found in Tax file'
                    })
                    
        except Exception as e:
            raise Exception(f"Error during validation: {str(e)}")
    
    def generate_report(self, output_path: Path):
        """Generate Excel report with results"""
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Create formats
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#D8E4BC',
                'border': 1
            })
            
            error_format = workbook.add_format({
                'bg_color': '#FFC7CE',
                'font_color': '#9C0006'
            })
            
            # Write summary sheet
            summary_data = {
                'Check Type': ['P/N Match', 'Description Match', 'HSN Match', 'Duty Match'],
                'Total Errors': [len(self.results[k]) for k in self.results.keys()]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            # Format summary sheet
            summary_sheet = writer.sheets['Summary']
            summary_sheet.set_column('A:B', 20)
            
            # Write detailed sheets
            for check_type, data in self.results.items():
                if data:  # Only create sheet if there are errors
                    df = pd.DataFrame(data)
                    sheet_name = check_type.upper()
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # Format sheet
                    sheet = writer.sheets[sheet_name]
                    sheet.set_column('A:D', 20)
                    
                    # Apply header format
                    for col_num, value in enumerate(df.columns.values):
                        sheet.write(0, col_num, value, header_format)

def main():
    """Main function to run the validator"""
    try:
        # Initialize validator
        validator = ExcelValidator()
        
        # Set up file paths (modify these as needed)
        check_file = Path("_Import CheckList-SIM1058324 New._CheckList_Data_17.43.45.xlsx")
        bom_file = Path("24HC01713海运物料 - 副本.xlsx")
        tax_file = Path("税率汇总-20230822.xlsx")
        output_file = Path("validation_report.xlsx")
        
        # Load and validate data
        df_check, df_bom, df_tax = validator.load_files(check_file, bom_file, tax_file)
        validator.validate_data(df_check, df_bom, df_tax)
        
        # Generate report
        validator.generate_report(output_file)
        print(f"Validation complete. Report generated at: {output_file}")
        
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main() 