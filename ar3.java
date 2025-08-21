

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
import os

def extract_section(df, start_keyword, stop_keywords):
    """Extracts values from a section starting with start_keyword until one of stop_keywords is hit"""
    collecting = False
    values = []
    for _, row in df.iterrows():
        for cell in row:
            if pd.isna(cell):
                continue
            text = str(cell).strip().upper()
            if start_keyword in text:
                collecting = True
                continue
            if any(kw in text for kw in stop_keywords) and collecting:
                return values
            if collecting:
                if "." in text or text.startswith("&&") or text.startswith("R1-"):
                    values.append(str(cell).strip())
    return values

def parse_and_transform_excel(input_file, output_file):
    sheet_dict = pd.read_excel(input_file, sheet_name=None, header=None)

    wb = Workbook()
    ws = wb.active
    ws.title = "All_Jobs"

    # Headers
    headers = ["SheetName", "Job ID", "Files IN", "Files OUT"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h).font = Font(bold=True)

    row_no = 2

    for sheet_name, df in sheet_dict.items():
        print(f"Processing sheet: {sheet_name}")

        # Try structured layout (FILES IN / FILES OUT sections)
        files_in = extract_section(df, "FILES IN", ["FILES OUT", "TABLES", "EXECUTION"])
        files_out = extract_section(df, "FILES OUT", ["TABLES", "EXECUTION", "BATCH DETAILS"])

        if files_in or files_out:
            job_id = sheet_name
            for f in files_in:
                ws.cell(row=row_no, column=1, value=sheet_name)
                ws.cell(row=row_no, column=2, value=job_id)
                ws.cell(row=row_no, column=3, value=f)
                ws.cell(row=row_no, column=4, value="")  # no OUT here
                row_no += 1
            for f in files_out:
                ws.cell(row=row_no, column=1, value=sheet_name)
                ws.cell(row=row_no, column=2, value=job_id)
                ws.cell(row=row_no, column=3, value="")  # no IN here
                ws.cell(row=row_no, column=4, value=f)
                row_no += 1
        else:
            # Fallback: mapping layout (Col A=Job, Col B=IN, Col C=OUT)
            for _, row in df.iterrows():
                job_id = str(row[0]).strip() if pd.notna(row[0]) else sheet_name
                file_in = str(row[1]).strip() if len(row) > 1 and pd.notna(row[1]) else ""
                file_out = str(row[2]).strip() if len(row) > 2 and pd.notna(row[2]) else ""

                # skip empty rows
                if not file_in and not file_out:
                    continue

                # avoid putting the same dataset in both IN and OUT
                if file_in and file_out and file_in == file_out:
                    file_out = ""

                ws.cell(row=row_no, column=1, value=sheet_name)
                ws.cell(row=row_no, column=2, value=job_id)
                ws.cell(row=row_no, column=3, value=file_in)
                ws.cell(row=row_no, column=4, value=file_out)
                row_no += 1

    wb.save(output_file)
    print(f"✅ Output saved to {output_file}")
