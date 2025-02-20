# Excel File Processing Implementation Logic

## 1. Data Structure Requirements
- Original Excel file with multiple invoices
- Output Excel file with:
  - Each invoice data in separate tabs
  - Tab names based on invoice numbers
  - Consistent column headers across all tabs

## 2. Processing Steps

### A. Initial Setup
1. Load the Excel file using pandas
2. Identify the header row (row containing "P/N", "Desc", "HSN", etc.)
3. Store these column headers for reuse in new tabs

### B. Data Preprocessing
1. Remove rows in yellow frame:
   - Identify and skip rows containing:
     - "Job No Si/M/"
     - "BL No. HASLC"
     - "Port Of Loading"
   - Skip any empty rows before the main table

### C. Invoice Detection and Data Extraction
1. Scan rows for invoice pattern:
   - Look for rows containing "Invoice:" pattern
   - Extract invoice number (e.g., "17/31" from "Invoice: 24HC01713-1S dt. 28-Dec-2024 Invoice 17/31")
   - Use regular expression: `Invoice:`

2. Data Processing Loop:
   ```pseudocode
   current_invoice = None
   current_data = []
   
   for each row in excel:
       if row matches invoice_pattern:
           if current_invoice exists:
               save_to_tab(current_data, current_invoice)
           current_invoice = extract_invoice_number(row)
           current_data = []
       else:
           if current_invoice exists:
               current_data.append(row)
   
   # Don't forget to save the last batch
   if current_data:
       save_to_tab(current_data, current_invoice)
   ```

### D. Tab Creation and Data Writing
1. For each invoice:
   - Create new tab named with invoice number
   - Write stored column headers
   - Write all rows until next invoice or end of file
   - Apply consistent formatting

## 3. Error Handling
- Check for missing or malformed invoice numbers
- Validate column headers consistency
- Handle empty data sections
- Verify tab name validity (remove invalid characters)

## 4. Output Validation
1. Verify all invoices are processed
2. Check data integrity in each tab
3. Confirm header consistency
4. Validate row counts

## 5. Enhancement Considerations
- Add progress tracking
- Include summary tab with invoice list
- Preserve original formatting
- Add error logging
- Support for multiple file processing
