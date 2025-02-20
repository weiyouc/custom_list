import pandas as pd
import re
from pathlib import Path
import logging
from typing import List, Dict, Tuple
import argparse

class ExcelConverter:
    def __init__(self):
        # Set up logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger(__name__)
        
        # Simplified pattern to match any row containing "Invoice:"
        self.invoice_pattern = re.compile(r'Invoice:', re.IGNORECASE)
        
        # Headers to skip (yellow frame content)
        self.skip_patterns = [
            'Job No SI/M',
            'BL No. HASLC',
            'Port Of Loading'
        ]
        
        # Updated patterns to handle trailing hyphens
        self.part_no_pattern = re.compile(r'PART NO\.?\s*([\w\.-]+?)(?:-\s*(?=MODEL NO|$)|-\s*$|(?=MODEL NO|$))', re.IGNORECASE)
        self.model_no_pattern = re.compile(r'MODEL NO\.?\s*([\w\.-]+?)(?:\s*$|-\s*$)', re.IGNORECASE)
        
        # Add pattern for +OR- replacement
        self.plus_minus_pattern = re.compile(r'\s*\+OR-\s*', re.IGNORECASE)
        
    def is_header_row(self, row: pd.Series) -> bool:
        """Check if row contains main column headers"""
        # Updated to match your exact column headers
        required_columns = ['P/N', 'Desc', 'HSN']  # Changed 'DESC' to 'Desc'
        row_str = ' '.join(str(val) for val in row.values)
        self.logger.debug(f"Checking header row: {row_str}")
        
        # More lenient header check
        matches = [col for col in required_columns if col.upper() in row_str.upper()]
        if matches:
            self.logger.info(f"Found header row with columns: {matches}")
            return True
        return False
    
    def is_invoice_row(self, row: pd.Series) -> Tuple[bool, str]:
        """
        Check if row contains invoice information and extract invoice number
        Returns: (is_invoice, invoice_number)
        """
        row_str = ' '.join(str(val) for val in row.values)
        self.logger.debug(f"Checking invoice row: {row_str}")
        if self.invoice_pattern.search(row_str):
            # Extract just the invoice number (e.g., "24HC01713-1S")
            match = re.search(r'(\d+HC\d+-\d+[A-Z]*)', row_str)
            if match:
                invoice_num = match.group(1)
                self.logger.info(f"Found invoice: {invoice_num}")
                return True, invoice_num
            return False, ''
        return False, ''
    
    def should_skip_row(self, row: pd.Series) -> bool:
        """Check if row should be skipped (yellow frame content)"""
        row_str = ' '.join(str(val) for val in row.values)
        should_skip = any(pattern in row_str for pattern in self.skip_patterns)
        if should_skip:
            self.logger.debug(f"Skipping row: {row_str}")
        return should_skip
    
    def clean_description(self, desc: str) -> str:
        """Clean description text by replacing +OR- with ± and other formatting"""
        desc = str(desc).strip()
        # Replace +OR- with ±
        desc = self.plus_minus_pattern.sub('±', desc)
        return desc
    
    def split_description(self, desc_str: str) -> Dict[str, str]:
        """Split description string into component parts"""
        # Convert to string and clean
        desc_str = str(desc_str).strip()
        
        # Extract Model No
        model_no_match = self.model_no_pattern.search(desc_str)
        
        # Get the base description (everything before PART NO)
        base_desc = desc_str.split('-PART NO')[0].strip()
        # Remove trailing hyphen if exists
        base_desc = re.sub(r'-\s*$', '', base_desc)
        
        # Clean the base description
        base_desc = self.clean_description(base_desc)
        
        return {
            'Description': base_desc,
            'Model No': model_no_match.group(1).strip() if model_no_match else ''
        }
    
    def process_dataframe(self, df: pd.DataFrame, header_row: pd.Series) -> pd.DataFrame:
        """Process the dataframe to split description column"""
        # Find the description column index and category column index
        desc_col = None
        category_col = None
        for idx, col in enumerate(header_row):
            if str(col).upper().strip() == 'DESC':
                desc_col = idx
            elif str(col) == 'Category':
                category_col = idx
        
        if desc_col is None:
            self.logger.warning("Description column not found")
            return df
        
        # Split descriptions
        split_data = []
        item_number = 1
        
        for _, row in df.iterrows():
            desc_parts = self.split_description(row[desc_col])
            
            # Create new row with Item number and split description
            new_row = [item_number]
            new_row.extend(list(row))
            
            # Update Description
            new_row[desc_col + 1] = desc_parts['Description']
            
            # Update Category/Item name with first part of Description
            if category_col is not None:
                item_name = desc_parts['Description'].split('-')[0].strip()
                new_row[category_col + 1] = item_name
            
            # Add Model No only
            new_row.extend([desc_parts['Model No']])
            split_data.append(new_row)
            
            item_number += 1
        
        # Create new header with Item Nos. column and additional columns
        new_header = ['Item Nos.']
        header_list = list(header_row)
        
        # Replace column names
        header_list = [
            'Description' if col == 'Desc' 
            else 'India HS code' if col == 'HSN'
            else 'Item name' if col == 'Category'
            else 'Quantity PCS' if col == 'Qty'
            else 'Unit Price USD' if col == 'Price'
            else 'Amount USD' if col == 'Value Amt'
            else col 
            for col in header_list
        ]
        
        new_header.extend(header_list)
        new_header.extend(['Model Nos.'])
        
        # Create new dataframe with split data
        new_df = pd.DataFrame(split_data, columns=new_header)
        
        # Round Amount USD column to 2 decimal places
        amount_col = 'Amount USD'
        if amount_col in new_df.columns:
            new_df[amount_col] = new_df[amount_col].round(2)
        
        # Reorder columns to match specified order - Update column names to match exactly
        first_columns = ['Item Nos.', 'Model Nos.', 'P/N', 'Description', 'Quantity PCS', 'Unit Price USD', 'Amount USD']
        other_columns = [col for col in new_df.columns if col not in first_columns]
        ordered_columns = first_columns + other_columns
        
        # Debug print to check available columns
        self.logger.debug(f"Available columns: {new_df.columns.tolist()}")
        self.logger.debug(f"Attempting to reorder columns: {ordered_columns}")
        
        # Reorder the DataFrame columns
        new_df = new_df[ordered_columns]
        
        return new_df
    
    def process_sheet(self, df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
        """Process a single sheet and return processed invoice data"""
        # Initialize variables
        header_row = None
        current_invoice = None
        current_data = []
        sheet_data = {}  # Dictionary to store data for each invoice
        
        # Process each row
        for idx, row in df.iterrows():
            self.logger.debug(f"Processing row {idx}")
            
            # Skip yellow frame content
            if self.should_skip_row(row):
                self.logger.info(f"--Skipping row: {' '.join(str(val) for val in row.values)}")
                continue
            
            # Check if this is the header row
            if header_row is None and self.is_header_row(row):
                header_row = row
                self.logger.info(f"Found header row at index {idx}: {header_row.tolist()}")
                continue
            
            # Check if this is an invoice row
            is_invoice, invoice_number = self.is_invoice_row(row)
            
            if is_invoice:
                # Save current batch if exists
                if current_invoice and current_data:
                    self.logger.info(f"Saving {len(current_data)} rows for invoice {current_invoice}")
                    sheet_data[current_invoice] = self.process_dataframe(pd.DataFrame(current_data), header_row)
                
                # Start new batch
                current_invoice = invoice_number
                current_data = []
                self.logger.info(f"Starting new invoice: {current_invoice}")
                
            elif current_invoice and header_row is not None:
                # Add row to current batch
                current_data.append(row)
                self.logger.debug(f"Added row to invoice {current_invoice}")
        
        # Save last batch
        if current_invoice and current_data:
            self.logger.info(f"Saving final {len(current_data)} rows for invoice {current_invoice}")
            sheet_data[current_invoice] = self.process_dataframe(pd.DataFrame(current_data), header_row)
        
        return sheet_data

    def process_excel(self, input_path: Path, output_path: Path):
        """Main function to process the Excel file"""
        try:
            # Read all sheets from the Excel file
            self.logger.info(f"Reading input file: {input_path}")
            excel_file = pd.ExcelFile(input_path)
            sheet_names = excel_file.sheet_names
            self.logger.info(f"Found {len(sheet_names)} sheets: {sheet_names}")
            
            # Process each sheet
            all_processed_data = {}
            
            for sheet_name in sheet_names:
                self.logger.info(f"Processing sheet: {sheet_name}")
                # Read the current sheet
                df = pd.read_excel(input_path, sheet_name=sheet_name, header=None)
                self.logger.info(f"Total rows in sheet: {len(df)}")
                
                # Process the sheet
                sheet_data = self.process_sheet(df)
                
                # Add processed data to overall results
                if sheet_data:
                    self.logger.info(f"Found {len(sheet_data)} invoices in sheet {sheet_name}")
                    all_processed_data.update(sheet_data)
                else:
                    self.logger.warning(f"No invoice data found in sheet: {sheet_name}")
            
            # Check if we found any data across all sheets
            if not all_processed_data:
                self.logger.error("No invoice data was found in any sheet!")
                return
            
            self.logger.info(f"Found total {len(all_processed_data)} invoices across all sheets")
            
            # Write to output file
            self.logger.info(f"Writing output file: {output_path}")
            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                # Write each invoice to its own sheet
                for invoice_num, data in all_processed_data.items():
                    sheet_name = re.sub(r'[\\/*\[\]:?]', '', invoice_num)
                    if len(sheet_name) > 31:
                        sheet_name = sheet_name[:31]
                    self.logger.info(f"Writing sheet {sheet_name} with {len(data)} rows")
                    
                    # Write to sheet
                    data.to_excel(
                        writer,
                        sheet_name=sheet_name,
                        index=False
                    )
                    
                    # Get the xlsxwriter worksheet object
                    worksheet = writer.sheets[sheet_name]
                    
                    # Format headers
                    header_format = writer.book.add_format({
                        'bold': True,
                        'bg_color': '#D8E4BC',
                        'border': 1
                    })
                    
                    # Apply header format and adjust column widths
                    for col_num, value in enumerate(data.columns):
                        worksheet.write(0, col_num, value, header_format)
                        max_length = max(
                            data[value].astype(str).apply(len).max(),
                            len(str(value))
                        )
                        worksheet.set_column(col_num, col_num, max_length + 2)
            
            self.logger.info("Processing completed successfully")
            
        except Exception as e:
            self.logger.error(f"Error processing file: {str(e)}")
            raise

def main():
    # Set up argument parser
    parser = argparse.ArgumentParser(description='Convert Excel invoice files to processed format')
    parser.add_argument('input_file', type=str, help='Path to the input Excel file')
    parser.add_argument('output_file', type=str, help='Path to save the processed Excel file')
    parser.add_argument('--debug', action='store_true', help='Enable debug logging')
    
    args = parser.parse_args()
    
    # Set debug level if requested
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    
    # Create and run converter
    converter = ExcelConverter()
    converter.process_excel(Path(args.input_file), Path(args.output_file))

if __name__ == "__main__":
    main()
