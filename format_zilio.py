import re

# --- CONFIGURATION ---
INPUT_FILE = "quran_zilio_final.txt"
OUTPUT_FILE = "quran_zilio_formatted_with_checks.txt"

def format_verses_smart():
    print(f"1. Reading {INPUT_FILE}...")
    
    with open(INPUT_FILE, "r", encoding="utf-8") as f:
        content = f.read()

    # --- STEP 1: GLOBAL CLEANUP ---
    # Fix hyphenation (word- \n break)
    content = re.sub(r'-\s*\n\s*', '', content)
    # Replace newlines with spaces to treat text as a stream
    content = content.replace('\n', ' ')
    # Remove double spaces
    content = re.sub(r'\s+', ' ', content).strip()

    # --- STEP 2: SPLIT BY NUMBERS ---
    # We split by digits. 
    # Result: ['', '1', ' text...', '2', ' text...', '1', ' text...']
    tokens = re.split(r'(\d+)', content)
    
    formatted_lines = []
    
    # State Variables
    current_surah = 1
    expected_verse = 1
    previous_verse_num = 0
    
    # Start loop from index 1 (where the first number should be)
    i = 1 
    
    print("2. Processing verses and checking for gaps...")

    while i < len(tokens) - 1:
        raw_num = tokens[i]
        text_part = tokens[i+1].strip()
        
        # Safety: Ensure it's actually a number
        if not raw_num.isdigit():
            i += 2
            continue
            
        found_num = int(raw_num)

        # --- LOGIC: IS THIS A NEW SURAH? ---
        # If the number dropped (e.g., was 286, now is 1 or 2), it's a new Surah.
        # We also check if it's the very first verse of the file (previous_verse_num == 0)
        if previous_verse_num > 0 and found_num < previous_verse_num:
            current_surah += 1
            expected_verse = 1 # Reset expectation
            # print(f"--- Detected New Surah {current_surah} (Verse reset to {found_num}) ---")

        # --- LOGIC: IS IT A NUMBER INSIDE THE TEXT? ---
        # If we jump from verse 5 to 1000, it's likely the number "1000" inside the text, not verse 1000.
        # Threshold: If jump is > 50 and we aren't at the start of a surah, assume it's text.
        if (found_num - expected_verse) > 50:
            # It's likely text. Append this number and text to the PREVIOUS line
            if formatted_lines:
                formatted_lines[-1] += f" {found_num} {text_part}"
            i += 2
            continue

        # --- LOGIC: GAP DETECTION (MISSING LINES) ---
        while expected_verse < found_num:
            # We expected 1, but found 2. Or expected 5, found 7.
            missing_ref = f"{current_surah}:{expected_verse}"
            print(f"!! MISSING VERSE DETECTED: {missing_ref}")
            
            # Write marker to file
            formatted_lines.append(f"{missing_ref} [MISSING LINE]")
            
            expected_verse += 1

        # --- LOGIC: NORMAL WRITE ---
        # If found_num matches expected_verse, we are good.
        final_line = f"{current_surah}:{found_num} {text_part}"
        formatted_lines.append(final_line)
        
        # Prepare for next iteration
        previous_verse_num = found_num
        expected_verse = found_num + 1
        i += 2

    # --- SAVE ---
    print(f"3. Saving to {OUTPUT_FILE}...")
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(formatted_lines))

    print("Done! Check the console output above for any missing verses.")

if __name__ == "__main__":
    format_verses_smart()