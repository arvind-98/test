



def extract_job_data_from_sheet(df, sheet_name):
    """
    Extract job data from a single sheet with better section detection
    Handles both structured (FILES IN/OUT headers) and mapping style sheets.
    """
    job_data = {
        'job_id': None,
        'files_in': [],
        'files_out': []
    }

    print(f"\nProcessing sheet: {sheet_name}")

    # -------- Case 1: Structured format (FILES IN / FILES OUT sections) --------
    files_in = find_files_in_section(df, "FILES IN")
    files_out = find_files_in_section(df, "FILES OUT")

    if files_in or files_out:
        # Try to detect job id from first 20 rows
        for idx in range(min(20, len(df))):
            for col in df.columns:
                if pd.notna(df.iloc[idx, col]):
                    val = str(df.iloc[idx, col]).strip()
                    if re.match(r'^[A-Z]{3,}[0-9]{3,4}$', val):  # e.g., ISCA0100
                        job_data['job_id'] = val
                        break
            if job_data['job_id']:
                break
        job_data['files_in'] = files_in
        job_data['files_out'] = files_out
        return job_data

    # -------- Case 2: Mapping style sheet (col A=Job, col B=IN, col C=OUT) --------
    print("Structured headers not found → assuming mapping style sheet")

    for idx, row in df.iterrows():
        job_id = str(row[0]).strip() if pd.notna(row[0]) else None
        if job_id and re.match(r'^[A-Z]{3,}[0-9]{3,4}$', job_id):
            job_data['job_id'] = job_id

        # Files IN in column B
        if len(row) > 1 and pd.notna(row[1]) and is_valid_filename(str(row[1]).strip()):
            job_data['files_in'].append(str(row[1]).strip())

        # Files OUT in column C
        if len(row) > 2 and pd.notna(row[2]) and is_valid_filename(str(row[2]).strip()):
            job_data['files_out'].append(str(row[2]).strip())

    # If job id not found, default to sheet name
    if not job_data['job_id']:
        job_data['job_id'] = sheet_name.upper()

    return job_data



import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import os
import re

def find_files_in_section(df, section_name):
    """
    Find files specifically in the 'Files IN' or 'Files OUT' section
    """
    files = []
    section_found = False
    section_end = False
    
    for idx, row in df.iterrows():
        # Check if we found the section header
        if not section_found:
            for col in df.columns:
                if pd.notna(df.iloc[idx, col]):
                    cell_value = str(df.iloc[idx, col]).strip().upper()
                    if section_name.upper() in cell_value:
                        section_found = True
                        print(f"Found section '{section_name}' at row {idx}")
                        break
            continue
        
        # If we're in the section, look for files
        if section_found and not section_end:
            row_files = []
            for col in df.columns:
                if pd.notna(df.iloc[idx, col]):
                    cell_value = str(df.iloc[idx, col]).strip()
                    
                    # Check if this is the start of another section
                    upper_cell = cell_value.upper()
                    if any(keyword in upper_cell for keyword in ['FILES OUT', 'FILES IN', 'TABLES', 'EXECUTION', 'BATCH DETAILS']) and section_name.upper() not in upper_cell:
                        section_end = True
                        break
                    
                    # Skip mapping or conversion columns
                    if 'MAPPED_TO' in cell_value.upper() or 'CONVERT REQUIRED' in cell_value.upper() or cell_value.upper() in ['Y/N/?', 'Y', 'N']:
                        continue
                    
                    # Check if this looks like a valid file name
                    if is_valid_filename(cell_value):
                        row_files.append(cell_value)
            
            # Add unique files from this row
            for file in row_files:
                if file not in files:
                    files.append(file)
                    print(f"Added {section_name} file: {file}")
            
            # If no files found in several consecutive rows, assume section ended
            if not row_files and len(files) > 0:
                empty_row_count = 0
                # Check next few rows for content
                for check_idx in range(idx, min(idx + 3, len(df))):
                    has_content = False
                    for col in df.columns:
                        if pd.notna(df.iloc[check_idx, col]) and str(df.iloc[check_idx, col]).strip():
                            has_content = True
                            break
                    if not has_content:
                        empty_row_count += 1
                
                if empty_row_count >= 2:  # Multiple empty rows suggest section end
                    section_end = True
    
    return files

def is_valid_filename(text):
    """
    More strict validation for file names to avoid false positives
    """
    text = text.strip()
    
    # Basic length check
    if len(text) < 3 or len(text) > 80:
        return False
    
    # Skip obvious non-file patterns
    skip_patterns = [
        'total', 'expected', 'accessed', 'source', 'start-end', 'seconds',
        'y/n/?', 'y', 'n', 'convert required', 'file path', 'file name',
        'file type', 'program name', 'utilities used', 'mapped_to',
        'description', 'comments', 'job steps', 'execution time'
    ]
    
    if text.lower() in skip_patterns or any(pattern in text.lower() for pattern in ['mapped_to', 'convert required']):
        return False
    
    # Must contain dots for dataset names or be temporary files
    if not ('.' in text or text.startswith('&&') or text.startswith('R1-')):
        return False
    
    # Valid file patterns
    file_patterns = [
        r'^[A-Z0-9&]+\.[A-Z0-9.()+-]+$',        # Standard dataset names
        r'^&&[A-Z0-9]+$',                        # Temporary datasets
        r'^R1-\.[A-Z0-9.]+$',                    # R1-.DATASET.NAME
        r'^[A-Z]+\.[A-Z0-9.()+-]+$',            # Prefix.dataset.name
    ]
    
    return any(re.match(pattern, text.upper()) for pattern in file_patterns)

def extract_job_data_from_sheet(df, sheet_name):
    """
    Extract job data from a single sheet with better section detection
    """
    job_data = {
        'job_id': None,
        'files_in': [],
        'files_out': []
    }
    
    print(f"\n{'='*50}")
    print(f"Processing sheet: {sheet_name}")
    print(f"Sheet dimensions: {df.shape}")
    
    # Find job ID - look for patterns like ISCA0100, ISCADRAP
    for idx in range(min(20, len(df))):  # Check first 20 rows
        for col in df.columns:
            if pd.notna(df.iloc[idx, col]):
                cell_value = str(df.iloc[idx, col]).strip()
                # Job ID patterns
                if re.match(r'^[A-Z]{4,}[0-9]{3,4}$|^[A-Z]{6,}[A-Z]*$', cell_value):
                    job_data['job_id'] = cell_value
                    print(f"Found Job ID: {cell_value}")
                    break
        if job_data['job_id']:
            break
    
    # Default job ID if not found
    if not job_data['job_id']:
        job_data['job_id'] = sheet_name.upper() if sheet_name else "UNKNOWN_JOB"
    
    # Extract Files IN and Files OUT separately
    job_data['files_in'] = find_files_in_section(df, 'FILES IN')
    job_data['files_out'] = find_files_in_section(df, 'FILES OUT')
    
    # If sections not found, try alternative extraction
    if not job_data['files_in'] and not job_data['files_out']:
        print("Section headers not found, trying pattern-based extraction...")
        all_files = extract_all_files_from_sheet(df)
        
        # Simple heuristic: assume first half are input, second half are output
        mid_point = len(all_files) // 2
        job_data['files_in'] = all_files[:mid_point]
        job_data['files_out'] = all_files[mid_point:]
    
    print(f"Final count - Files IN: {len(job_data['files_in'])}, Files OUT: {len(job_data['files_out'])}")
    return job_data

def extract_all_files_from_sheet(df):
    """
    Extract all valid file names from the sheet as fallback
    """
    files = []
    for idx, row in df.iterrows():
        for col in df.columns:
            if pd.notna(df.iloc[idx, col]):
                cell_value = str(df.iloc[idx, col]).strip()
                if is_valid_filename(cell_value) and cell_value not in files:
                    files.append(cell_value)
    return files

def create_clean_output_excel(all_jobs_data, output_file):
    """
    Create a clean single Excel file with Files IN and Files OUT in separate columns
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Job_Files_Analysis"
    
    # Formatting
    job_id_font = Font(bold=True, color="FF0000")  # Red
    header_font = Font(bold=True, color="000000")  # Black
    
    # Headers - Files IN in column B, Files OUT in column C
    ws['A1'] = 'Job ID'
    ws['B1'] = 'Files IN' 
    ws['C1'] = 'Files OUT'
    
    for col in ['A1', 'B1', 'C1']:
        ws[col].font = header_font
    
    current_row = 2
    
    for job_data in all_jobs_data:
        job_id = job_data['job_id']
        files_in = job_data['files_in']
        files_out = job_data['files_out']
        
        # Job ID in column A (red) - only on the first row of each job
        ws.cell(row=current_row, column=1, value=job_id)
        ws.cell(row=current_row, column=1).font = job_id_font
        
        # Files IN in column B - each file in its own row
        files_in_start_row = current_row
        for i, file_name in enumerate(files_in):
            ws.cell(row=files_in_start_row + i, column=2, value=file_name)
        
        # Files OUT in column C - each file in its own row, starting from same row as Job ID
        files_out_start_row = current_row  
        for i, file_name in enumerate(files_out):
            ws.cell(row=files_out_start_row + i, column=3, value=file_name)
        
        # Move to next job section - ensure enough space for both IN and OUT files
        max_files = max(len(files_in), len(files_out), 1)  # At least 1 for the job ID row
        current_row += max_files + 2  # Add spacing between jobs
        
        print(f"Added job '{job_id}': {len(files_in)} Files IN (col B), {len(files_out)} Files OUT (col C)")
    
    # Column widths
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 50  # Files IN
    ws.column_dimensions['C'].width = 50  # Files OUT
    
    wb.save(output_file)
    print(f"\nOutput saved to: {output_file}")
    print("Layout: Column A = Job ID, Column B = Files IN, Column C = Files OUT")

def parse_and_transform_excel(input_file, output_file):
    """
    Main parsing function
    """
    try:
        print(f"Reading input file: {input_file}")
        
        # Read all sheets
        sheet_dict = pd.read_excel(input_file, sheet_name=None, header=None)
        print(f"Found {len(sheet_dict)} sheet(s)")
        
        all_jobs_data = []
        
        # Process each sheet
        for sheet_name, df in sheet_dict.items():
            job_data = extract_job_data_from_sheet(df, sheet_name)
            
            # Only include if we found files
            if job_data['files_in'] or job_data['files_out']:
                all_jobs_data.append(job_data)
                print(f"✓ Sheet '{sheet_name}' processed successfully")
            else:
                print(f"✗ Sheet '{sheet_name}' - no valid files found")
        
        if not all_jobs_data:
            print("ERROR: No valid job data found in any sheet!")
            return
        
        # Create output
        create_clean_output_excel(all_jobs_data, output_file)
        
        # Summary
        print(f"\n{'='*60}")
        print(f"SUCCESS - Transformation completed!")
        print(f"{'='*60}")
        print(f"📁 Input:  {input_file}")
        print(f"📄 Output: {output_file} (Single sheet)")
        print(f"🔢 Jobs processed: {len(all_jobs_data)}")
        
        for job_data in all_jobs_data:
            print(f"   • {job_data['job_id']}: {len(job_data['files_in'])} Files IN → {len(job_data['files_out'])} Files OUT")
        
    except Exception as e:
        print(f"❌ ERROR: {e}")
        import traceback
        traceback.print_exc()

def main():
    """
    Interactive main function
    """
    print("🔄 Excel Job Files Parser")
    print("=" * 30)
    
    # Get file paths
    input_file = input("📥 Input Excel file (default: input_data.xlsx): ").strip()
    if not input_file:
        input_file = "input_data.xlsx"
    
    output_file = input("📤 Output Excel file (default: job_files_output.xlsx): ").strip()  
    if not output_file:
        output_file = "job_files_output.xlsx"
    
    # Validate input
    if not os.path.exists(input_file):
        print(f"❌ File not found: {input_file}")
        return
    
    # Process
    parse_and_transform_excel(input_file, output_file)

if __name__ == "__main__":
    main()


import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import os
import re

def find_files_in_section(df, section_name):
    """
    Find files specifically in the 'Files IN' or 'Files OUT' section
    """
    files = []
    section_found = False
    section_end = False
    
    for idx, row in df.iterrows():
        # Check if we found the section header
        if not section_found:
            for col in df.columns:
                if pd.notna(df.iloc[idx, col]):
                    cell_value = str(df.iloc[idx, col]).strip().upper()
                    if section_name.upper() in cell_value:
                        section_found = True
                        print(f"Found section '{section_name}' at row {idx}")
                        break
            continue
        
        # If we're in the section, look for files
        if section_found and not section_end:
            row_files = []
            for col in df.columns:
                if pd.notna(df.iloc[idx, col]):
                    cell_value = str(df.iloc[idx, col]).strip()
                    
                    # Check if this is the start of another section
                    upper_cell = cell_value.upper()
                    if any(keyword in upper_cell for keyword in ['FILES OUT', 'FILES IN', 'TABLES', 'EXECUTION', 'BATCH DETAILS']) and section_name.upper() not in upper_cell:
                        section_end = True
                        break
                    
                    # Skip mapping or conversion columns
                    if 'MAPPED_TO' in cell_value.upper() or 'CONVERT REQUIRED' in cell_value.upper() or cell_value.upper() in ['Y/N/?', 'Y', 'N']:
                        continue
                    
                    # Check if this looks like a valid file name
                    if is_valid_filename(cell_value):
                        row_files.append(cell_value)
            
            # Add unique files from this row
            for file in row_files:
                if file not in files:
                    files.append(file)
                    print(f"Added {section_name} file: {file}")
            
            # If no files found in several consecutive rows, assume section ended
            if not row_files and len(files) > 0:
                empty_row_count = 0
                # Check next few rows for content
                for check_idx in range(idx, min(idx + 3, len(df))):
                    has_content = False
                    for col in df.columns:
                        if pd.notna(df.iloc[check_idx, col]) and str(df.iloc[check_idx, col]).strip():
                            has_content = True
                            break
                    if not has_content:
                        empty_row_count += 1
                
                if empty_row_count >= 2:  # Multiple empty rows suggest section end
                    section_end = True
    
    return files

def is_valid_filename(text):
    """
    More strict validation for file names to avoid false positives
    """
    text = text.strip()
    
    # Basic length check
    if len(text) < 3 or len(text) > 80:
        return False
    
    # Skip obvious non-file patterns
    skip_patterns = [
        'total', 'expected', 'accessed', 'source', 'start-end', 'seconds',
        'y/n/?', 'y', 'n', 'convert required', 'file path', 'file name',
        'file type', 'program name', 'utilities used', 'mapped_to',
        'description', 'comments', 'job steps', 'execution time'
    ]
    
    if text.lower() in skip_patterns or any(pattern in text.lower() for pattern in ['mapped_to', 'convert required']):
        return False
    
    # Must contain dots for dataset names or be temporary files
    if not ('.' in text or text.startswith('&&') or text.startswith('R1-')):
        return False
    
    # Valid file patterns
    file_patterns = [
        r'^[A-Z0-9&]+\.[A-Z0-9.()+-]+$',        # Standard dataset names
        r'^&&[A-Z0-9]+$',                        # Temporary datasets
        r'^R1-\.[A-Z0-9.]+$',                    # R1-.DATASET.NAME
        r'^[A-Z]+\.[A-Z0-9.()+-]+$',            # Prefix.dataset.name
    ]
    
    return any(re.match(pattern, text.upper()) for pattern in file_patterns)

def extract_job_data_from_sheet(df, sheet_name):
    """
    Extract job data from a single sheet with better section detection
    """
    job_data = {
        'job_id': None,
        'files_in': [],
        'files_out': []
    }
    
    print(f"\n{'='*50}")
    print(f"Processing sheet: {sheet_name}")
    print(f"Sheet dimensions: {df.shape}")
    
    # Find job ID - look for patterns like ISCA0100, ISCADRAP
    for idx in range(min(20, len(df))):  # Check first 20 rows
        for col in df.columns:
            if pd.notna(df.iloc[idx, col]):
                cell_value = str(df.iloc[idx, col]).strip()
                # Job ID patterns
                if re.match(r'^[A-Z]{4,}[0-9]{3,4}$|^[A-Z]{6,}[A-Z]*$', cell_value):
                    job_data['job_id'] = cell_value
                    print(f"Found Job ID: {cell_value}")
                    break
        if job_data['job_id']:
            break
    
    # Default job ID if not found
    if not job_data['job_id']:
        job_data['job_id'] = sheet_name.upper() if sheet_name else "UNKNOWN_JOB"
    
    # Extract Files IN and Files OUT separately
    job_data['files_in'] = find_files_in_section(df, 'FILES IN')
    job_data['files_out'] = find_files_in_section(df, 'FILES OUT')
    
    # If sections not found, try alternative extraction
    if not job_data['files_in'] and not job_data['files_out']:
        print("Section headers not found, trying pattern-based extraction...")
        all_files = extract_all_files_from_sheet(df)
        
        # Simple heuristic: assume first half are input, second half are output
        mid_point = len(all_files) // 2
        job_data['files_in'] = all_files[:mid_point]
        job_data['files_out'] = all_files[mid_point:]
    
    print(f"Final count - Files IN: {len(job_data['files_in'])}, Files OUT: {len(job_data['files_out'])}")
    return job_data

def extract_all_files_from_sheet(df):
    """
    Extract all valid file names from the sheet as fallback
    """
    files = []
    for idx, row in df.iterrows():
        for col in df.columns:
            if pd.notna(df.iloc[idx, col]):
                cell_value = str(df.iloc[idx, col]).strip()
                if is_valid_filename(cell_value) and cell_value not in files:
                    files.append(cell_value)
    return files

def create_clean_output_excel(all_jobs_data, output_file):
    """
    Create a clean single Excel file with Files IN and Files OUT in separate columns
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Job_Files_Analysis"
    
    # Formatting
    job_id_font = Font(bold=True, color="FF0000")  # Red
    header_font = Font(bold=True, color="000000")  # Black
    
    # Headers - Files IN in column B, Files OUT in column C
    ws['A1'] = 'Job ID'
    ws['B1'] = 'Files IN' 
    ws['C1'] = 'Files OUT'
    
    for col in ['A1', 'B1', 'C1']:
        ws[col].font = header_font
    
    current_row = 2
    
    for job_data in all_jobs_data:
        job_id = job_data['job_id']
        files_in = job_data['files_in']
        files_out = job_data['files_out']
        
        # Job ID in column A (red) - only on the first row of each job
        ws.cell(row=current_row, column=1, value=job_id)
        ws.cell(row=current_row, column=1).font = job_id_font
        
        # Files IN in column B - each file in its own row
        files_in_start_row = current_row
        for i, file_name in enumerate(files_in):
            ws.cell(row=files_in_start_row + i, column=2, value=file_name)
        
        # Files OUT in column C - each file in its own row, starting from same row as Job ID
        files_out_start_row = current_row  
        for i, file_name in enumerate(files_out):
            ws.cell(row=files_out_start_row + i, column=3, value=file_name)
        
        # Move to next job section - ensure enough space for both IN and OUT files
        max_files = max(len(files_in), len(files_out), 1)  # At least 1 for the job ID row
        current_row += max_files + 2  # Add spacing between jobs
        
        print(f"Added job '{job_id}': {len(files_in)} Files IN (col B), {len(files_out)} Files OUT (col C)")
    
    # Column widths
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 50  # Files IN
    ws.column_dimensions['C'].width = 50  # Files OUT
    
    wb.save(output_file)
    print(f"\nOutput saved to: {output_file}")
    print("Layout: Column A = Job ID, Column B = Files IN, Column C = Files OUT")

def parse_and_transform_excel(input_file, output_file):
    """
    Main parsing function
    """
    try:
        print(f"Reading input file: {input_file}")
        
        # Read all sheets
        sheet_dict = pd.read_excel(input_file, sheet_name=None, header=None)
        print(f"Found {len(sheet_dict)} sheet(s)")
        
        all_jobs_data = []
        
        # Process each sheet
        for sheet_name, df in sheet_dict.items():
            job_data = extract_job_data_from_sheet(df, sheet_name)
            
            # Only include if we found files
            if job_data['files_in'] or job_data['files_out']:
                all_jobs_data.append(job_data)
                print(f"✓ Sheet '{sheet_name}' processed successfully")
            else:
                print(f"✗ Sheet '{sheet_name}' - no valid files found")
        
        if not all_jobs_data:
            print("ERROR: No valid job data found in any sheet!")
            return
        
        # Create output
        create_clean_output_excel(all_jobs_data, output_file)
        
        # Summary
        print(f"\n{'='*60}")
        print(f"SUCCESS - Transformation completed!")
        print(f"{'='*60}")
        print(f"📁 Input:  {input_file}")
        print(f"📄 Output: {output_file} (Single sheet)")
        print(f"🔢 Jobs processed: {len(all_jobs_data)}")
        
        for job_data in all_jobs_data:
            print(f"   • {job_data['job_id']}: {len(job_data['files_in'])} Files IN → {len(job_data['files_out'])} Files OUT")
        
    except Exception as e:
        print(f"❌ ERROR: {e}")
        import traceback
        traceback.print_exc()

def main():
    """
    Interactive main function
    """
    print("🔄 Excel Job Files Parser")
    print("=" * 30)
    
    # Get file paths
    input_file = input("📥 Input Excel file (default: input_data.xlsx): ").strip()
    if not input_file:
        input_file = "input_data.xlsx"
    
    output_file = input("📤 Output Excel file (default: job_files_output.xlsx): ").strip()  
    if not output_file:
        output_file = "job_files_output.xlsx"
    
    # Validate input
    if not os.path.exists(input_file):
        print(f"❌ File not found: {input_file}")
        return
    
    # Process
    parse_and_transform_excel(input_file, output_file)

if __name__ == "__main__":
    main()



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
        job_data['files_in'] = extract_file_list(df, files_in_start + 1, files_out_start)
    else:
        # If no explicit "Files IN" section, look for file patterns in the sheet
        job_data['files_in'] = extract_file_patterns(df, 'IN')
    
    # Extract Files OUT
    if files_out_start is not None:
        job_data['files_out'] = extract_file_list(df, files_out_start + 1, None)
    else:
        # If no explicit "Files OUT" section, look for file patterns in the sheet
        job_data['files_out'] = extract_file_patterns(df, 'OUT')
    
    # Remove duplicates while preserving order
    job_data['files_in'] = list(dict.fromkeys(job_data['files_in']))
    job_data['files_out'] = list(dict.fromkeys(job_data['files_out']))
    
    # Remove files that appear in both IN and OUT from IN list to avoid duplication
    job_data['files_in'] = [f for f in job_data['files_in'] if f not in job_data['files_out']]
    
    print(f"Files IN found: {len(job_data['files_in'])}")
    print(f"Files OUT found: {len(job_data['files_out'])}")
    
    return job_data

def extract_file_list(df, start_row, end_row):
    """
    Extract file names between start_row and end_row
    """
    files = []
    max_row = end_row if end_row is not None else len(df)
    max_rows_to_check = min(50, max_row - start_row)  # Prevent infinite loops
    
    for idx in range(start_row, min(start_row + max_rows_to_check, max_row)):
        if idx >= len(df):
            break
            
        row_has_files = False
        for col in range(len(df.columns)):
            if pd.notna(df.iloc[idx, col]):
                cell_value = str(df.iloc[idx, col]).strip()
                
                # Skip empty cells and common headers
                if (cell_value == '' or cell_value.lower() in ['nan', 'convert required', 'y/n/?', 'file path', 'file name', 'file type']):
                    continue
                
                # Stop if we hit another section
                if any(keyword in cell_value.upper() for keyword in ['FILES OUT', 'FILES IN', 'TABLES', 'EXECUTION', 'BATCH DETAILS']):
                    return files
                
                # Check if this looks like a file name
                if is_likely_filename(cell_value):
                    if cell_value not in files:  # Avoid duplicates
                        files.append(cell_value)
                        row_has_files = True
                        print(f"Found file: {cell_value}")
    
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
                    'file type', 'program name', 'utilities used', 'mapped_to']
    
    if text.lower() in skip_patterns or 'mapped_to' in text.lower():
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

def create_single_output_excel(all_jobs_data, output_file):
    """
    Create a single Excel file with one sheet containing all job data
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Job_Data"
    
    # Set up formatting
    job_id_font = Font(bold=True, color="FF0000")  # Red font
    header_font = Font(bold=True, color="000000")  # Black bold font
    
    current_row = 1
    
    # Process each job
    for job_idx, job_data in enumerate(all_jobs_data):
        job_id = job_data['job_id']
        files_in = job_data['files_in']
        files_out = job_data['files_out']
        
        # Add some spacing between jobs (except for the first job)
        if job_idx > 0:
            current_row += 2
        
        # Add job ID in red
        ws.cell(row=current_row, column=1, value=job_id)
        ws.cell(row=current_row, column=1).font = job_id_font
        current_row += 1
        
        # Find the maximum number of files to align properly
        max_files = max(len(files_in), len(files_out))
        
        # Add Files IN in column B
        for idx, file_name in enumerate(files_in):
            ws.cell(row=current_row + idx, column=2, value=file_name)
        
        # Add Files OUT in column D (separate column)
        for idx, file_name in enumerate(files_out):
            ws.cell(row=current_row + idx, column=4, value=file_name)
        
        # Move to next section
        current_row += max_files + 1
        
        print(f"Added job '{job_id}' with {len(files_in)} input files and {len(files_out)} output files")
    
    # Add column headers at the top for clarity
    ws.insert_rows(1)
    ws.cell(row=1, column=1, value="Job ID")
    ws.cell(row=1, column=2, value="Files IN")
    ws.cell(row=1, column=4, value="Files OUT")
    
    # Format headers
    for col in [1, 2, 4]:
        ws.cell(row=1, column=col).font = header_font
    
    # Adjust column widths
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 5   # Spacer column
    ws.column_dimensions['D'].width = 40
    
    # Save the workbook
    wb.save(output_file)
    print(f"\nSuccessfully created single output file: {output_file}")

def parse_and_transform_excel(input_file, output_file):
    """
    Main function to parse input Excel and create transformed output
    """
    try:
        # Read all sheets from the input file
        sheet_dict = pd.read_excel(input_file, sheet_name=None, header=None)
        print(f"Found {len(sheet_dict)} sheets in the input file")
        
        all_jobs_data = []
        
        # Process each sheet
        for sheet_name, df in sheet_dict.items():
            print(f"\n{'='*50}")
            print(f"Processing Sheet: {sheet_name}")
            print(f"{'='*50}")
            
            job_data = extract_job_data_from_sheet(df, sheet_name)
            
            # Only add if we found some data
            if job_data['files_in'] or job_data['files_out']:
                all_jobs_data.append(job_data)
            else:
                print(f"No file data found in sheet '{sheet_name}', skipping...")
        
        if not all_jobs_data:
            print("No valid data found in any sheet!")
            return
        
        # Create single output Excel with all jobs
        create_single_output_excel(all_jobs_data, output_file)
        
        # Print summary
        print(f"\n{'='*60}")
        print(f"TRANSFORMATION SUMMARY")
        print(f"{'='*60}")
        print(f"Input file: {input_file}")
        print(f"Output file: {output_file}")
        print(f"Total jobs processed: {len(all_jobs_data)}")
        print(f"Output: Single sheet with all job data")
        
        for job_data in all_jobs_data:
            print(f"  - {job_data['job_id']}: {len(job_data['files_in'])} Files IN, {len(job_data['files_out'])} Files OUT")
            
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
