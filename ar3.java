import openpyxl
import pandas as pd

# Load workbook
wb = openpyxl.load_workbook("CSP_Batch_Tracker.xlsm", data_only=True)

# List to store extracted rows
rows = []

# Loop through all sheets
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    
    # Find the header row (looking for "Files IN")
    for row in ws.iter_rows(min_row=1, max_row=100, values_only=True):
        if row and "Files IN" in str(row[0]):
            start_row = ws.iter_rows(min_row=row[0].row+1, values_only=True)
            break
    
    # Extract File IN (col A) and File OUT (col D)
    for r in ws.iter_rows(min_row=53, max_row=200, values_only=True):
        if not r[0] and not r[3]:
            continue
        file_in = r[0]
        file_out = r[3]
        if file_in or file_out:
            rows.append([sheet_name, file_in, file_out])

# Convert to DataFrame
df = pd.DataFrame(rows, columns=["Tab Name", "File IN", "File OUT"])

# Save to new Excel
df.to_excel("Extracted_FileMapping.xlsx", index=False)

print("Extraction completed. File saved as Extracted_FileMapping.xlsx")
