import tempfile
import subprocess
import re
import os

station_prefix = "st02wvaw"

def extract_station_names(text):
    station_set = set()

    # Remove space between digits like "st02wvaw0 123"
    compressed = re.sub(r'(?<=\d)\s+(?=\d)', '', text)
    print("\n[DEBUG] Compressed input:\n" + compressed)

    # Extract full station names like st02wvaw0123
    full_matches = re.findall(rf'{station_prefix}\d{{4}}', compressed, re.IGNORECASE)
    print(f"\n[DEBUG] Full matches found: {full_matches}")
    station_set.update(map(str.lower, full_matches))

    # Extract numeric suffixes (1-4 digits) — avoid duplicate if already in full form
    suffix_matches = re.findall(r'\b(\d{1,4})\b', compressed)
    print(f"[DEBUG] Suffix-only matches found: {suffix_matches}")
    for suffix in suffix_matches:
        padded = suffix.zfill(4)
        full_name = station_prefix + padded
        station_set.add(full_name)

    return sorted(station_set)

def main():
    # Create a temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode='w', encoding='utf-8') as tmp:
        temp_path = tmp.name
        tmp.write("Paste your ticket titles here.\nSave and close Notepad when done.\n")

    print(f"\n[INFO] Opening Notepad: {temp_path}")
    proc = subprocess.Popen(["notepad.exe", temp_path])
    proc.wait()

    # Read contents
    with open(temp_path, 'r', encoding='utf-8') as f:
        content = f.read()

    os.remove(temp_path)

    print("\n[INFO] Raw input from file:\n" + content)

    # Extract and show station names
    stations = extract_station_names(content)
    if stations:
        print("\n✅ Extracted station names:")
        for s in stations:
            print(s)
    else:
        print("\n⚠️ No station names found. Check your input format.")

    # Wait for user to press Enter before exiting
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()
