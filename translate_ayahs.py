import openpyxl
import os

# --- CONFIGURATION ---
EXCEL_FILE = 'vocab.xlsx'                        # Your Excel file
TXT_SOURCE = 'quran_zilio_formatted_with_checks.txt' # The clean text file we just made
SHEET_NAME = "Vocab with Examples"
REF_HEADER = "Ayahref"                           # Column A (Input)
TARGET_HEADER = "Meaning & Translation in Italian" # Column B (Output)

def automate_quran_translation():
    # 1. Build the Lookup Dictionary from the TXT file
    print(f"1. Reading translation from '{TXT_SOURCE}'...")
    quran_map = {}
    
    if not os.path.exists(TXT_SOURCE):
        print(f"Error: The source text file '{TXT_SOURCE}' was not found.")
        return

    with open(TXT_SOURCE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line: continue
            
            # The format is: "2:14 The Text is Here"
            # We split ONLY at the first space
            parts = line.split(" ", 1)
            
            if len(parts) == 2:
                key = parts[0].strip()   # "2:14"
                text = parts[1].strip()  # "The Text is Here"
                quran_map[key] = text
            elif len(parts) == 1:
                # Handle cases like "3:1 [MISSING LINE]" if there is no text after
                key = parts[0].strip()
                quran_map[key] = "[TEXT MISSING IN SOURCE]"

    print(f"   Database built. Loaded {len(quran_map)} verses.")

    # 2. Open the Excel File
    print(f"2. Opening Excel file '{EXCEL_FILE}'...")
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
    except FileNotFoundError:
        print(f"Error: Excel file '{EXCEL_FILE}' not found.")
        return

    if SHEET_NAME not in wb.sheetnames:
        print(f"Error: Sheet '{SHEET_NAME}' not found.")
        return
    
    ws = wb[SHEET_NAME]

    # 3. Find Column Indices
    ref_col_index = None
    target_col_index = None
    
    # Scan headers (Row 1)
    for cell in ws[1]:
        if cell.value == REF_HEADER:
            ref_col_index = cell.column
        elif cell.value == TARGET_HEADER:
            target_col_index = cell.column

    if ref_col_index is None:
        print(f"Error: Column '{REF_HEADER}' not found.")
        return

    # Create target column if it doesn't exist
    if target_col_index is None:
        print(f"Column '{TARGET_HEADER}' not found. Creating it...")
        target_col_index = ws.max_column + 1
        ws.cell(row=1, column=target_col_index).value = TARGET_HEADER

    # 4. Process Rows
    print("3. Filling translations...")
    updates_count = 0
    
    # Iterate through rows
    for row in range(2, ws.max_row + 1):
        ref_cell = ws.cell(row=row, column=ref_col_index)
        target_cell = ws.cell(row=row, column=target_col_index)
        
        raw_ref = ref_cell.value
        
        if raw_ref:
            # CLEANING: " 2 : 102 " -> "2:102"
            clean_ref = str(raw_ref).replace(" ", "").replace(":", ":").strip()
            
            # Lookup
            translation = quran_map.get(clean_ref)
            
            if translation:
                target_cell.value = translation
                updates_count += 1
            else:
                # Optional: Leave blank or write "Not Found"
                # target_cell.value = "Ref Not Found"
                pass

    # 5. Save
    print(f"4. Saving updates to '{EXCEL_FILE}'...")
    try:
        wb.save(EXCEL_FILE)
        print(f"Done! Updated {updates_count} rows.")
    except PermissionError:
        print("ERROR: Please close the Excel file before running this script!")

if __name__ == "__main__":
    automate_quran_translation()