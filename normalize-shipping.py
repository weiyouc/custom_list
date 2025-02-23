import pandas as pd
import re
from pathlib import Path
import warnings

# Add this to ignore pandas warnings too
pd.options.mode.chained_assignment = None

# Add this at the top to suppress openpyxl warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def normalize_shipping_file(input_file: str, output_file: str):
    """
    Process shipping list file to:
    1. Remove PL tab
    2. Clean invoice tabs to only keep shipping content
    3. Maintain required columns
    """
    # Load all sheets except PL tab
    all_sheets = pd.read_excel(input_file, sheet_name=None)
    sheets_to_process = {name: df for name, df in all_sheets.items() 
                        if not re.search(r'\bPL\b', name, flags=re.IGNORECASE)}
    
    processed_sheets = {}
    
    for sheet_name, df in sheets_to_process.items():
        # Step 1: Find the shipping content table
        shipping_df = extract_shipping_table(df)
        
        # Step 2: Keep only required columns
        filtered_df = filter_columns(shipping_df)
        
        if not filtered_df.empty:
            processed_sheets[sheet_name] = filtered_df
    
    # Save to new Excel file
    with pd.ExcelWriter(output_file) as writer:
        for sheet_name, df in processed_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"Processed file saved to: {output_file}")

def extract_shipping_table(df: pd.DataFrame) -> pd.DataFrame:
    """Find and extract the shipping content table from a sheet"""
    # Find the first row that looks like a table header
    header_row = find_header_row(df)
    if header_row is None:
        return pd.DataFrame()
    
    # Extract table data
    shipping_df = df.iloc[header_row:].copy()
    shipping_df.columns = clean_headers(shipping_df.iloc[0])
    
    # Remove empty rows and reset index
    cleaned_df = shipping_df.iloc[1:].dropna(how='all').reset_index(drop=True)
    
    # Add validation for Item No. column (using the standardized name)
    if 'Item No.' in cleaned_df.columns:
        # Keep only rows where Item No. looks valid (number or specific pattern)
        cleaned_df = cleaned_df[
            cleaned_df['Item No.'].astype(str).str.match(r'^\s*\d+|Item\s+No', na=False)
        ]
    
    return cleaned_df

def find_header_row(df: pd.DataFrame) -> int:
    """Find the first row containing shipping table headers"""
    # Use both variations to find the header
    required_columns = ['Item Nos.', 'Item No.', 'Model No.', 'P/N', 'Quantity PCS', 'Unit Price USD', 'Amount USD']
    
    for idx, row in df.iterrows():
        clean_row = [str(cell).strip() for cell in row.values]
        row_str = ' '.join(clean_row).lower()
        
        match_count = sum(1 for col in required_columns 
                         if col.lower() in row_str)
        
        if match_count >= 3:  # Found header row
            return idx
            
    return None

def get_header_variants(standard_name: str) -> list:
    """Get all known header variations for a column"""
    column_map = {
        'Item No.': ['Item No.', 'Item Number', 'Item Nos', 'Item', 'Item#', 'Item Code'],
        'Model No.': ['Model No.', 'Model Number', 'Model Nos'],
        'P/N': ['P/N', 'Part Number', 'Part No'],
        'Quantity PCS': ['Quantity PCS', 'QTY', 'Quantity'],
    }
    return column_map.get(standard_name, [standard_name])

def is_empty_row(row: pd.Series) -> bool:
    """Check if a row is essentially empty"""
    return row.dropna().empty

def clean_column_name(name: str) -> str:
    """Clean column name by removing newlines and extra spaces"""
    return str(name).replace('\n', '').replace('\r', '').strip()

def clean_headers(headers) -> list:
    """Normalize column headers"""
    column_map = {
        'Item No.': ['Item Nos.', 'Item Nos', 'Item Number', 'Item', 'Item#', 'Item Code'],
        'Model No.': ['Model No.', 'Model Number', 'Model Nos'],
        'P/N': ['P/N', 'Part Number', 'Part No'],
        'Description': ['Description', 'Desc'],
        'Quantity PCS': ['Quantity PCS', 'QTY', 'Quantity'],
        'Unit Price USD': ['Unit Price USD', 'Unit Price', 'Price'],
        'Amount USD': ['Amount USD', 'Amount', 'Total']
    }
    
    cleaned = []
    for header in headers:
        # Clean the header first
        header_str = clean_column_name(header)
        matched = False
        for standard_name, variants in column_map.items():
            # Compare cleaned versions
            if any(header_str.lower() == clean_column_name(v).lower() for v in variants):
                cleaned.append(standard_name)
                matched = True
                break
        if not matched:
            cleaned.append(header_str)
    return cleaned

def filter_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Keep only required columns"""
    required = ['Item No.', 'Model No.', 'P/N', 'Description', 
               'Quantity PCS', 'Unit Price USD', 'Amount USD']
    
    # Clean column names in DataFrame
    df.columns = [clean_column_name(col) for col in df.columns]
    
    # Debug print
    print("Available columns after cleaning:", df.columns.tolist())
    
    # Get columns that exist in the DataFrame
    cols_to_keep = []
    for col in required:
        if col in df.columns or clean_column_name(col) in df.columns:
            cols_to_keep.append(col)
    
    if not cols_to_keep:
        print("Warning: No required columns found in DataFrame")
        return pd.DataFrame()
    
    print("Keeping columns:", cols_to_keep)
    return df[cols_to_keep]

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='Normalize shipping list Excel file')
    parser.add_argument('input_file', help='Path to input shipping list Excel file')
    parser.add_argument('output_file', nargs='?', default=None,
                      help='Path for normalized output file (default: input path with _normalized suffix)')
    
    args = parser.parse_args()
    
    # Set default output path if not provided
    if not args.output_file:
        input_path = Path(args.input_file)
        args.output_file = input_path.parent / f"{input_path.stem}_normalized.xlsx"
    
    normalize_shipping_file(args.input_file, args.output_file)
