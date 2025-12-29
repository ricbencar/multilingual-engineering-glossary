"""
==============================================================================
SCRIPT: get_fonts.py
DESCRIPTION:
    Robust Noto Font Downloader.
    
    STRATEGY:
    - Bypasses Google Fonts API (which is blocking us).
    - Uses "Brute Force" discovery: It tries 3 different known GitHub folder 
      structures for every font until it finds the file.
    - Uses 'raw.githubusercontent.com' which allows automated downloads.
    - Automatically fetches special "Mega-Merge" fonts for extended Latin.
    - Fetches the massive Noto Sans CJK (super-font) for East Asian languages.

    USAGE:
    python get_fonts.py
==============================================================================
"""

import os
import requests

# ==============================================================================
# CONFIGURATION
# ==============================================================================

OUTPUT_DIR = "fonts"

# The list of files we need.
# We also provide the "Google Folder Name" (lowercase) for the backup mirror.
FONTS_NEEDED = [
    # (Filename, Google_Folder_Name_Lowercase)
    
    # --- 1. CORE FONTS (Latin/Cyrillic & Headers) ---
    ("NotoSans-Regular.ttf",           "notosans"),
    ("NotoSans-Bold.ttf",              "notosans"),  # REQUIRED for PDF Headers
    ("NotoSansLiving-Regular.ttf",     "notosans"),  # "Mega-Merge" for Turkish/Vietnamese
    
    # --- 2. SOUTH ASIAN (Indic) ---
    ("NotoSansDevanagari-Regular.ttf", "notosansdevanagari"),
    ("NotoSansGujarati-Regular.ttf",   "notosansgujarati"),
    ("NotoSansGurmukhi-Regular.ttf",   "notosansgurmukhi"),
    ("NotoSansBengali-Regular.ttf",    "notosansbengali"),
    ("NotoSansTamil-Regular.ttf",      "notosanstamil"),
    ("NotoSansTelugu-Regular.ttf",     "notosanstelugu"),
    
    # --- 3. SOUTHEAST ASIAN ---
    ("NotoSansThai-Regular.ttf",       "notosansthai"),
    ("NotoSansJavanese-Regular.ttf",   "notosansjavanese"),
    
    # --- 4. MIDDLE EASTERN (RTL) ---
    ("NotoSansArabic-Regular.ttf",     "notosansarabic"),
    ("NotoNastaliqUrdu-Regular.ttf",   "notonastaliqurdu"),
]

# CJK Configuration (Noto Sans CJK - Variable Font Collection)
# This file contains Simplified, Traditional, Japanese, and Korean in one binary.
CJK_FILENAME = "NotoSansCJK.ttc"
CJK_URL = "https://raw.githubusercontent.com/notofonts/noto-cjk/main/Sans/Variable/OTC/NotoSansCJK-VF.otf.ttc"

# ==============================================================================
# FUNCTIONS
# ==============================================================================

def get_candidate_urls(filename, google_folder):
    """
    Generates a list of possible URLs for a given font file.
    Different Noto fonts live in different repos/folder structures.
    """
    font_name_folder = filename.replace("-Regular.ttf", "").replace("-Bold.ttf", "")
    
    urls = []
    
    # Priority 0: Special Case for "Mega-Merge" (NotoSansLiving)
    if "Living" in filename:
        urls.append(f"https://raw.githubusercontent.com/notofonts/notofonts.github.io/main/megamerge/{filename}")
        return urls

    # Priority 1: Main Noto Repo (Hinted) - Best for Windows
    urls.append(f"https://raw.githubusercontent.com/notofonts/noto-fonts/main/hinted/ttf/{font_name_folder}/{filename}")
    
    # Priority 2: Main Noto Repo (Unhinted) - Backup
    urls.append(f"https://raw.githubusercontent.com/notofonts/noto-fonts/main/unhinted/ttf/{font_name_folder}/{filename}")
    
    # Priority 3: Latin-Greek-Cyrillic Repo (Specific for NotoSans/NotoSerif)
    urls.append(f"https://raw.githubusercontent.com/notofonts/latin-greek-cyrillic/main/hinted/ttf/{font_name_folder}/{filename}")
    urls.append(f"https://raw.githubusercontent.com/notofonts/latin-greek-cyrillic/main/unhinted/ttf/{font_name_folder}/{filename}")

    # Priority 4: Google Fonts Static Mirror (Modern Structure)
    urls.append(f"https://raw.githubusercontent.com/google/fonts/main/ofl/{google_folder}/static/{filename}")
    
    # Priority 5: Google Fonts Static Mirror (Old Structure)
    urls.append(f"https://raw.githubusercontent.com/google/fonts/main/ofl/{google_folder}/{filename}")

    return urls

def download_font_smart(filename, google_folder, output_dir):
    dest_path = os.path.join(output_dir, filename)
    
    if os.path.exists(dest_path):
        print(f"  [SKIP] {filename} exists.")
        return

    print(f"  [LOOKUP] Finding source for {filename}...")
    
    candidates = get_candidate_urls(filename, google_folder)
    
    for i, url in enumerate(candidates):
        try:
            # Short timeout, we just want to see if it exists
            r = requests.get(url, stream=True, timeout=10)
            
            if r.status_code == 200:
                print(f"    -> Found at Source #{i+1}. Downloading...")
                with open(dest_path, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)
                print("    -> Success.")
                return
            
        except Exception:
            pass # Just try the next one
            
    print(f"  [FAILURE] Could not find {filename} in any known repo.")

def download_cjk(output_dir):
    dest_path = os.path.join(output_dir, CJK_FILENAME)
    
    if os.path.exists(dest_path):
        print(f"  [SKIP] {CJK_FILENAME} exists.")
        return

    print(f"  [FETCH] {CJK_FILENAME} (Large File)...")
    try:
        r = requests.get(CJK_URL, stream=True, timeout=120) # Increased timeout for large file
        if r.status_code == 200:
            with open(dest_path, 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
            print("    -> Success.")
        else:
            print(f"    [ERROR] CJK Download failed with code {r.status_code}")
    except Exception as e:
        print(f"    [ERROR] {e}")

# ==============================================================================
# MAIN
# ==============================================================================

def main():
    print("--- Starting Brute-Force Font Download ---")
    
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        print(f"Phase 1: Created '{OUTPUT_DIR}'")

    print("\nPhase 2: Scanning Repositories...")
    for fname, gfolder in FONTS_NEEDED:
        download_font_smart(fname, gfolder, OUTPUT_DIR)

    print("\nPhase 3: CJK Download (Noto Sans CJK)...")
    download_cjk(OUTPUT_DIR)

    # Count
    if os.path.exists(OUTPUT_DIR):
        count = len([f for f in os.listdir(OUTPUT_DIR) if f.endswith(('.ttf', '.ttc'))])
    else:
        count = 0
        
    print("\n" + "="*50)
    print(f"DONE. Total files in '{OUTPUT_DIR}': {count}")
    print("="*50)

if __name__ == "__main__":
    main()