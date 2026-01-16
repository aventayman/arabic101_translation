import requests
import openpyxl

# --- CONFIGURATION ---
FILE_NAME = 'vocab.xlsx'        # The name of your existing Excel file
SHEET_NAME = "Vocab with Examples"
REF_HEADER = "Ayahref"          # The header of the reference column
TARGET_HEADER = "Meaning & Translation in Italian" # The header where translation goes

def automate_quran_translation():
    # 1. Fetch the Translation Data
    print("1. Fetching Hamza Roberto Piccardo translation...")
    url = "http://api.alquran.cloud/v1/quran/it.piccardo"
    try:
        response = requests.get(url)
        response.raise_for_status()
    except Exception as e:
        print(f"Error fetching API: {e}")
        return

    # 2. Build the Lookup Dictionary
    print("2. Building database...")
    quran_map = {}
    data = response.json()['data']['surahs']
    for surah in data:
        surah_num = str(surah['number'])
        for ayah in surah['ayahs']:
            ayah_num = str(ayah['numberInSurah'])
            # We store the key as "2:14" (clean, no spaces)
            key = f"{surah_num}:{ayah_num}"
            quran_map[key] = ayah['text']

    # 3. Open the Excel File (Preserving Formatting)
    print(f"3. Opening '{FILE_NAME}'...")
    try:
        wb = openpyxl.load_workbook(FILE_NAME)
    except FileNotFoundError:
        print(f"Error: File '{FILE_NAME}' not found.")
        return

    if SHEET_NAME not in wb.sheetnames:
        print(f"Error: Sheet '{SHEET_NAME}' not found.")
        return
    
    ws = wb[SHEET_NAME]

    # 4. Find the Column Indices
    ref_col_index = None
    target_col_index = None
    
    # We scan the first row (headers) to find where our columns are
    header_row = ws[1] 
    max_col = ws.max_column

    for cell in header_row:
        if cell.value == REF_HEADER:
            ref_col_index = cell.column
        elif cell.value == TARGET_HEADER:
            target_col_index = cell.column

    # Error if reference column is missing
    if ref_col_index is None:
        print(f"Error: Could not find column header '{REF_HEADER}'")
        return

    # If target column doesn't exist, create the header at the end
    if target_col_index is None:
        print(f"Column '{TARGET_HEADER}' not found. Creating it...")
        target_col_index = max_col + 1
        ws.cell(row=1, column=target_col_index).value = TARGET_HEADER

    # 5. Process the Rows
    print("4. Updating rows...")
    # Iterate from row 2 (skipping header) to the last row
    for row in range(2, ws.max_row + 1):
        ref_cell = ws.cell(row=row, column=ref_col_index)
        target_cell = ws.cell(row=row, column=target_col_index)
        
        raw_ref = ref_cell.value
        
        if raw_ref:
            # CLEANING STEP: Convert to string and remove ALL spaces
            # " 2 : 102 " becomes "2:102"
            clean_ref = str(raw_ref).replace(" ", "").strip()
            
            # Lookup translation
            translation = quran_map.get(clean_ref)
            
            if translation:
                target_cell.value = translation
            else:
                # Optional: Mark not found if you want, or leave blank
                target_cell.value = "Ref Not Found"

    # 6. Save the file
    print("5. Saving file...")
    try:
        wb.save(FILE_NAME)
        print("Done! Formatting preserved.")
    except PermissionError:
        print("Error: Please close the Excel file before running the script!")

if __name__ == "__main__":
    automate_quran_translation()