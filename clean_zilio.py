import pdfplumber
import re

# --- CONFIGURATION ---
PDF_PATH = "IL+CORANO.pdf"
OUTPUT_FILE = "quran_zilio_final.txt"
START_PAGE = 75
END_PAGE = 492

def is_sura_header(line):
    # Detects "SURA X", "S U R A X", "SURA I", "Su¯ra"
    if len(line) > 50: return False
    # This saves lines like "sopra di loro, 9..."
    if "," in line: return False
    # Check for digits
    if re.search(r'S.*[uū].*R.*A.*\d+', line, re.IGNORECASE):
        return True
    # Roman Numerals check (Sura I, Sura II) - also requiring 'U'
    if re.search(r'S.*[uū].*R.*A.*[IVX]+$', line, re.IGNORECASE):
        return True
    return False

def is_basmala(line):
    # Checks for "Nel nome di Dio" (case insensitive)
    return "nel nome di dio, il clemente, il compassionevole" in line.lower()

def clean_and_extract():
    print(f"1. Scanning PDF Pages {START_PAGE}-{END_PAGE}...")
    
    # --- PHASE 1: Raw Extraction (Remove Noise) ---
    pre_processed_lines = []
    
    with pdfplumber.open(PDF_PATH) as pdf:
        for i in range(START_PAGE - 1, END_PAGE):
            try:
                page = pdf.pages[i]
                text = page.extract_text()
            except IndexError:
                break 

            if not text: continue
            
            raw_lines = text.split('\n')
            for line in raw_lines:
                line = line.strip()
                if not line: continue 

                # Remove Page Numbers (Just digits)
                if re.fullmatch(r'\d+', line):
                    continue
                # Remove "IL CORANO" Header
                if re.fullmatch(r'(\d+\s*)?IL CORANO(\s*\d+)?', line, re.IGNORECASE):
                    continue

                pre_processed_lines.append(line)
            
            if (i+1) % 50 == 0:
                print(f"Extracted page {i+1}...")

    print(f"Total lines extracted: {len(pre_processed_lines)}")
    print("2. Applying Logic: Headers, Titles, and Basmalas...")

    # --- PHASE 2: Structural Logic ---
    final_output = []
    found_first_header = False # Flag to track Surah 1
    
    i = 0
    while i < len(pre_processed_lines):
        line = pre_processed_lines[i]
        
        # Check if line is a Sura Header
        if is_sura_header(line):
            # Look ahead to see what follows (Title? Basmala?)
            # We grab the next 2 lines safely
            line_after_1 = pre_processed_lines[i+1] if i+1 < len(pre_processed_lines) else ""
            line_after_2 = pre_processed_lines[i+2] if i+2 < len(pre_processed_lines) else ""
            
            basmala_is_in_slot_2 = is_basmala(line_after_2)

            # --- CASE A: FIRST SURAH (Al-Fatiha) ---
            if not found_first_header:
                found_first_header = True
                if basmala_is_in_slot_2:
                    # Header (i) -> Title (i+1) -> Basmala/Verse1 (i+2)
                    # User: "Keep the first basmala". "Remove title if basmala follows".
                    # ACTION: Skip Header. Skip Title. KEEP Basmala.
                    i += 2 
                    continue
            
            # --- CASE B: STANDARD SURAH (with Basmala) ---
            if basmala_is_in_slot_2:
                # Header (i) -> Title (i+1) -> Basmala (i+2)
                # User: "Remove title if basmala follows".
                # ACTION: Skip Header. Skip Title. Skip Basmala.
                i += 3
                continue
            
            # --- CASE C: SURAH WITHOUT BASMALA (e.g. At-Tawbah) ---
            else:
                # Header (i) -> Title (i+1) -> Verse 1 (i+2 is NOT Basmala)
                # User: "Remove line after header ONLY if basmala follows."
                # Condition failed.
                # ACTION: Skip Header ONLY. Keep Title. Keep Verse 1.
                i += 1
                continue

        # --- NORMAL LINE PROCESSING ---
        
        # 1. Fix "Stuck Numbers" ("1Nel" -> "1 Nel")
        line = re.sub(r'(\d)([A-Za-zÀ-ÿ])', r'\1 \2', line)

        # 2. Add Blank Line for New Surah (Verse 1 detection)
        # "I want to add a blank line only when the numbers go back to one"
        if line.startswith("1 ") or line == "1":
            if final_output: # Avoid blank line at very top
                final_output.append("")

        final_output.append(line)
        i += 1

    # --- SAVE ---
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(final_output))
    
    print(f"\nDone! Saved to {OUTPUT_FILE}")

if __name__ == "__main__":
    clean_and_extract()