
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import os
import re

def extract_job_data_from_sheet(df, sheet_name):
    """
    Extract job data from a single sheet
    Returns: dict with job_id, files_in, files_out
    """
    job_data = {
        'job_id': None,
        'files_in': [],
        'files_out': []
    }
    
    print(f"\nProcessing sheet: {sheet_name}")
    print(f"Sheet dimensions: {df.shape}")
    
    # Find job ID (pattern: ISCA followed by numbers, or other job patterns)
    for idx, row in df.iterrows():
        for col in df.columns:
            if pd.notna(df.iloc[idx, col]):
                cell_value = str(df.iloc[idx, col]).strip()
                # Look for job ID patterns (ISCA0100, ISCADRAP, etc.)
                if re.match(r'^[A-Z]{3,}[0-9]{2,}$|^[A-Z]{4,}[A-Z]*$', cell_value):
                    job_data['job_id'] = cell_value
                    print(f"Found Job ID: {cell_value}")
                    break
        if job_data['job_id']:
            break
    
    # If no job ID found, use sheet name or create one
    if not job_data['job_id']:
        job_data['job_id'] = sheet_name.upper() if sheet_name else "JOB001"
        print(f"No Job ID found, using: {job_data['job_id']}")
    
    # Find Files IN and Files OUT sections
    files_in_start = None
    files_out_start = None
    
    for idx, row in df.iterrows():
        for col in df.columns:
            if pd.notna(df.iloc[idx, col]):
                cell_value = str(df.iloc[idx, col]).strip().upper()
                if 'FILES IN' in cell_value or 'INPUT FILES' in cell_value:
                    files_in_start = idx
                    print(f"Found 'Files IN' at row: {idx}")
                elif 'FILES OUT' in cell_value or 'OUTPUT FILES' in cell_value:
                    files_out_start = idx
                    print(f"Found 'Files OUT' at row: {idx}")
    
    # Extract Files IN
    if files_in_start is not None:
        job_data['files_in'] = extract_file_list(df, files_in_start + 1, 'IN')
    else:
        # If no explicit "Files IN" section, look for file patterns in the sheet
        job_data['files_in'] = extract_file_patterns(df, 'IN')
    
    # Extract Files OUT
    if files_out_start is not None:
        job_data['files_out'] = extract_file_list(df, files_out_start + 1, 'OUT')
    else:
        # If no explicit "Files OUT" section, look for file patterns in the sheet
        job_data['files_out'] = extract_file_patterns(df, 'OUT')
    
    print(f"Files IN found: {len(job_data['files_in'])}")
    print(f"Files OUT found: {len(job_data['files_out'])}")
    
    return job_data

def extract_file_list(df, start_row, file_type):
    """
    Extract file names starting from a specific row
    """
    files = []
    max_rows_to_check = 50  # Prevent infinite loops
    
    for idx in range(start_row, min(start_row + max_rows_to_check, len(df))):
        if idx >= len(df):
            break
            
        row_has_files = False
        for col in range(len(df.columns)):
            if pd.notna(df.iloc[idx, col]):
                cell_value = str(df.iloc[idx, col]).strip()
                
                # Skip empty cells and common headers
                if (cell_value == '' or cell_value.lower() in ['nan', 'convert required', 'y/n/?', 'file path', 'file name', 'file type']):
                    continue
                
                # Check if this looks like a file name
                if is_likely_filename(cell_value):
                    files.append(cell_value)
                    row_has_files = True
                    print(f"Found {file_type} file: {cell_value}")
        
        # If we hit several empty rows, stop looking
        if not row_has_files and idx > start_row + 5:
            # Check if we've hit another section
            for col in range(len(df.columns)):
                if pd.notna(df.iloc[idx, col]):
                    cell_value = str(df.iloc[idx, col]).strip().upper()
                    if any(keyword in cell_value for keyword in ['FILES OUT', 'FILES IN', 'TABLES', 'EXECUTION', 'BATCH']):
                        return files
    
    return files

def extract_file_patterns(df, file_type):
    """
    Extract file names by looking for common file name patterns across the entire sheet
    """
    files = []
    
    for idx, row in df.iterrows():
        for col in df.columns:
            if pd.notna(df.iloc[idx, col]):
                cell_value = str(df.iloc[idx, col]).strip()
                
                if is_likely_filename(cell_value):
                    # Avoid duplicates
                    if cell_value not in files:
                        files.append(cell_value)
    
    print(f"Extracted {len(files)} potential files using pattern matching for {file_type}")
    return files

def is_likely_filename(text):
    """
    Determine if a text string looks like a filename
    """
    text = text.strip()
    
    # Skip if too short or too long
    if len(text) < 3 or len(text) > 80:
        return False
    
    # Skip common non-file values
    skip_patterns = ['total', 'expected', 'accessed', 'source', 'start-end', 'seconds', 
                    'y/n/?', 'y', 'n', 'convert required', 'file path', 'file name', 
                    'file type', 'program name', 'utilities used']
    
    if text.lower() in skip_patterns:
        return False
    
    # Common file patterns
    file_patterns = [
        r'^[A-Z0-9]+\.[A-Z0-9]+\.[A-Z0-9.]+$',  # DATASET.NAME.PATTERN
        r'^[A-Z0-9&]+\.[A-Z0-9.]*$',            # Simple dataset names
        r'^&&[A-Z0-9]+$',                       # Temporary datasets
        r'^[A-Z0-9]+\.[A-Z0-9]+\.[A-Z0-9]+\([+\-0-9]+\)$',  # Datasets with generation
        r'^[A-Z][0-9]+-\.[A-Z0-9.]+$',         # R1-.DATASET.NAME
        r'^[A-Z]{2,}\.[A-Z0-9.]+$',            # Common prefixes
    ]
    
    for pattern in file_patterns:
        if re.match(pattern, text.upper()):
            return True
    
    # Additional checks for file-like strings
    if ('.' in text and 
        len(text.split('.')) >= 2 and 
        all(part.replace('(', '').replace(')', '').replace('+', '').replace('-', '').replace('0', '').replace('1', '').isalnum() 
            for part in text.split('.') if part)):
        return True
    
    return False

def create_output_excel(job_data_list, output_file):
    """
    Create the output Excel file with multiple sheets based on job data
    """
    wb = Workbook()
    
    # Remove the default sheet
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    
    for i, job_data in enumerate(job_data_list):
        job_id = job_data['job_id']
        files_in = job_data['files_in']
        files_out = job_data['files_out']
        
        # Create a new sheet with job_id as the name
        ws = wb.create_sheet(title=job_id)
        
        # Set up formatting
        job_id_font = Font(bold=True, color="FF0000")  # Red font
        
        # Add job ID in red (like the screenshot)
        ws['A1'] = job_id
        ws['A1'].font = job_id_font
        
        # Set up column headers (if needed, make them subtle)
        current_row = 2
        
        # Create the layout similar to the screenshot
        max_files = max(len(files_in), len(files_out))
        
        # Add Files IN in column B
        for idx, file_name in enumerate(files_in):
            ws.cell(row=current_row + idx, column=2, value=file_name)
        
        # Add Files OUT in column C (or D based on your preference)
        out_column = 4 if len(files_in) > 0 else 3  # Adjust spacing based on content
        for idx, file_name in enumerate(files_out):
            ws.cell(row=current_row + idx, column=out_column, value=file_name)
        
        # If you want to add some mapping or relationship between IN and OUT files
        if files_in and files_out:
            # Add a sample mapping in column C for the first file
            ws.cell(row=current_row, column=3, value=f"MAPPED_TO_{job_id}")
        
        # Adjust column widths
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 40
        ws.column_dimensions['C'].width = 40
        ws.column_dimensions['D'].width = 40
        
        print(f"Created sheet '{job_id}' with {len(files_in)} input files and {len(files_out)} output files")
    
    # Save the workbook
    wb.save(output_file)
    print(f"\nSuccessfully created output file: {output_file}")

def parse_and_transform_excel(input_file, output_file):
    """
    Main function to parse input Excel and create transformed output
    """
    try:
        # Read all sheets from the input file
        sheet_dict = pd.read_excel(input_file, sheet_name=None, header=None)
        print(f"Found {len(sheet_dict)} sheets in the input file")
        
        job_data_list = []
        
        # Process each sheet
        for sheet_name, df in sheet_dict.items():
            print(f"\n{'='*50}")
            print(f"Processing Sheet: {sheet_name}")
            print(f"{'='*50}")
            
            job_data = extract_job_data_from_sheet(df, sheet_name)
            
            # Only add if we found some data
            if job_data['files_in'] or job_data['files_out']:
                job_data_list.append(job_data)
            else:
                print(f"No file data found in sheet '{sheet_name}', skipping...")
        
        if not job_data_list:
            print("No valid data found in any sheet!")
            return
        
        # Create output Excel
        create_output_excel(job_data_list, output_file)
        
        # Print summary
        print(f"\n{'='*60}")
        print(f"TRANSFORMATION SUMMARY")
        print(f"{'='*60}")
        print(f"Input file: {input_file}")
        print(f"Output file: {output_file}")
        print(f"Sheets processed: {len(job_data_list)}")
        
        for job_data in job_data_list:
            print(f"  - {job_data['job_id']}: {len(job_data['files_in'])} IN, {len(job_data['files_out'])} OUT")
            
    except Exception as e:
        print(f"Error processing Excel file: {e}")
        import traceback
        traceback.print_exc()

def main():
    """
    Main function with user input
    """
    print("Excel Data Parser and Transformer")
    print("=" * 40)
    
    # Get input file from user
    input_file = input("Enter the input Excel file path (or press Enter for 'input_data.xlsx'): ").strip()
    if not input_file:
        input_file = "input_data.xlsx"
    
    # Get output file from user
    output_file = input("Enter the output Excel file path (or press Enter for 'transformed_output.xlsx'): ").strip()
    if not output_file:
        output_file = "transformed_output.xlsx"
    
    # Check if input file exists
    if not os.path.exists(input_file):
        print(f"\nError: Input file '{input_file}' not found!")
        print("Please make sure the file exists and try again.")
        return
    
    print(f"\nStarting transformation...")
    print(f"Input: {input_file}")
    print(f"Output: {output_file}")
    
    # Run the transformation
    parse_and_transform_excel(input_file, output_file)

# Direct usage function
def transform_excel_files(input_path, output_path):
    """
    Direct function call for programmatic usage
    """
    parse_and_transform_excel(input_path, output_path)

if __name__ == "__main__":
    main()





imimport pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import os
import re

def extract_job_data_from_sheet(df, sheet_name):
    """
    Extract job data from a single sheet
    Returns: dict with job_id, files_in, files_out
    """
    job_data = {
        'job_id': None,
        'files_in': [],
        'files_out': []
    }
    
    print(f"\nProcessing sheet: {sheet_name}")
    print(f"Sheet dimensions: {df.shape}")
    
    # Find job ID (pattern: ISCA followed by numbers, or other job patterns)
    for idx, row in df.iterrows():
        for col in df.columns:
            if pd.notna(df.iloc[idx, col]):
                cell_value = str(df.iloc[idx, col]).strip()
                # Look for job ID patterns (ISCA0100, ISCADRAP, etc.)
                if re.match(r'^[A-Z]{3,}[0-9]{2,}$|^[A-Z]{4,}[A-Z]*$', cell_value):
                    job_data['job_id'] = cell_value
                    print(f"Found Job ID: {cell_value}")
                    break
        if job_data['job_id']:
            break
    
    # If no job ID found, use sheet name or create one
    if not job_data['job_id']:
        job_data['job_id'] = sheet_name.upper() if sheet_name else "JOB001"
        print(f"No Job ID found, using: {job_data['job_id']}")
    
    # Find Files IN and Files OUT sections
    files_in_start = None
    files_out_start = None
    
    for idx, row in df.iterrows():
        for col in df.columns:
            if pd.notna(df.iloc[idx, col]):
                cell_value = str(df.iloc[idx, col]).strip().upper()
                if 'FILES IN' in cell_value or 'INPUT FILES' in cell_value:
                    files_in_start = idx
                    print(f"Found 'Files IN' at row: {idx}")
                elif 'FILES OUT' in cell_value or 'OUTPUT FILES' in cell_value:
                    files_out_start = idx
                    print(f"Found 'Files OUT' at row: {idx}")
    
    # Extract Files IN
    if files_in_start is not None:
        job_data['files_in'] = extract_file_list(df, files_in_start + 1, 'IN')
    else:
        # If no explicit "Files IN" section, look for file patterns in the sheet
        job_data['files_in'] = extract_file_patterns(df, 'IN')
    
    # Extract Files OUT
    if files_out_start is not None:
        job_data['files_out'] = extract_file_list(df, files_out_start + 1, 'OUT')
    else:
        # If no explicit "Files OUT" section, look for file patterns in the sheet
        job_data['files_out'] = extract_file_patterns(df, 'OUT')
    
    print(f"Files IN found: {len(job_data['files_in'])}")
    print(f"Files OUT found: {len(job_data['files_out'])}")
    
    return job_data

def extract_file_list(df, start_row, file_type):
    """
    Extract file names starting from a specific row
    """
    files = []
    max_rows_to_check = 50  # Prevent infinite loops
    
    for idx in range(start_row, min(start_row + max_rows_to_check, len(df))):
        if idx >= len(df):
            break
            
        row_has_files = False
        for col in range(len(df.columns)):
            if pd.notna(df.iloc[idx, col]):
                cell_value = str(df.iloc[idx, col]).strip()
                
                # Skip empty cells and common headers
                if (cell_value == '' or cell_value.lower() in ['nan', 'convert required', 'y/n/?', 'file path', 'file name', 'file type']):
                    continue
                
                # Check if this looks like a file name
                if is_likely_filename(cell_value):
                    files.append(cell_value)
                    row_has_files = True
                    print(f"Found {file_type} file: {cell_value}")
        
        # If we hit several empty rows, stop looking
        if not row_has_files and idx > start_row + 5:
            # Check if we've hit another section
            for col in range(len(df.columns)):
                if pd.notna(df.iloc[idx, col]):
                    cell_value = str(df.iloc[idx, col]).strip().upper()
                    if any(keyword in cell_value for keyword in ['FILES OUT', 'FILES IN', 'TABLES', 'EXECUTION', 'BATCH']):
                        return files
    
    return files

def extract_file_patterns(df, file_type):
    """
    Extract file names by looking for common file name patterns across the entire 
