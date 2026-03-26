import urllib.request
import json
import re
import urllib.parse
import os

# --- Configuration ---
REPO_API_URL = "https://api.github.com/repos/one-pace/one-pace-public-subtitles/git/trees/main?recursive=1"
# NEW: Using the blazing fast jsDelivr CDN!
RAW_BASE_URL = "https://cdn.jsdelivr.net/gh/one-pace/one-pace-public-subtitles@main/"
OUTPUT_JSON = "meta/subtitles.json"

# Map the folder names to our Stremio ID prefixes
ARC_MAP = {
    "romancedawn": "RO", "orangetown": "OR", "syrupvillage": "SY", "gaimon": "GA",
    "baratie": "BA", "arlongpark": "AR", "buggyscrewadventure": "BUGGYS_CREW",
    "loguetown": "LO", "reversemountain": "RM", "whiskypeak": "WH", "littlegarden": "LI",
    "drumisland": "DI", "arabasta": "AL", "alabasta": "AL", "jaya": "JA", "skypiea": "SK",
    "longringlongland": "LR", "waterseven": "WS", "enieslobby": "EN", "postenieslobby": "PEN",
    "thrillerbark": "TB", "sabaodyarchipelago": "SAB", "amazonlily": "AM", "impeldown": "IM",
    "marineford": "MA", "postwar": "PW", "returntosabaody": "RTS", "fishmanisland": "FI",
    "punkhazard": "PH", "dressrosa": "DR", "zou": "ZO", "wholecakeisland": "WC",
    "reverie": "REV", "wano": "WA", "egghead": "EH"
}

# Map file languages to Stremio's standard 3-letter ISO codes
LANG_MAP = {
    "ar": "ara", "cs": "cze", "cz": "cze", "de": "ger", "en": "eng", 
    "en cc": "eng", "en-cc": "eng", "en_cc": "eng", # Added hyphen/underscore support
    "es": "spa", "fi": "fin", "fr": "fre", "he": "heb", "it": "ita", "pl": "pol",
    "pt": "por", "ptbr": "por", "pt-br": "por", "pt_br": "por", # Added pt-br variations
    "ru": "rus", "tr": "tur", "id": "ind", "nl": "dut", "vi": "vie",
    "ja": "jpn", "typesetting": "eng"
}

def clean_string(s):
    # Removes numbers and spaces to match ARC_MAP perfectly
    return re.sub(r'[\d\s\-]', '', s).lower()

def main():
    print("[?] Fetching subtitle repository tree from GitHub...")
    req = urllib.request.Request(REPO_API_URL, headers={'User-Agent': 'Mozilla/5.0'})
    
    try:
        with urllib.request.urlopen(req) as response:
            tree_data = json.loads(response.read().decode())
    except Exception as e:
        print(f"[-] Failed to fetch from GitHub API: {e}")
        return

    subtitles_dict = {}

    for item in tree_data.get("tree", []):
        path = item.get("path", "")
        if not path.endswith(".ass"):
            continue

        parts = path.split("/")
        if len(parts) < 3:
            continue
        
        filename = parts[-1]        # e.g., "fi 17 es.ass"
        ep_folder = parts[-2]       # e.g., "17"
        arc_folder = parts[-3]      # e.g., "26 Fishman Island"

        arc_key = clean_string(arc_folder)
        prefix = ARC_MAP.get(arc_key)

        if not prefix:
            continue

        try:
            ep_num = int(ep_folder)
        except ValueError:
            continue

        stremio_id = f"{prefix}_{ep_num}"

        # Extract language string from the filename 
        name_without_ext = filename.rsplit('.', 1)[0]
        ep_str = str(ep_num).zfill(2)
        
        idx = name_without_ext.rfind(ep_str)
        if idx != -1:
            raw_lang_str = name_without_ext[idx + len(ep_str):].strip()
        else:
            raw_lang_str = "" # Fallback if we can't find the episode number in the filename

        # Find the actual 3-letter language code
        lang_code = "eng" # Default to English if not found
        # NEW: use .lower() here so "He" becomes "he" and successfully matches "heb"!
        for word in raw_lang_str.lower().split():
            if word in LANG_MAP:
                lang_code = LANG_MAP[word]
                break 
                
        # Generate a unique ID
        safe_lang_suffix = raw_lang_str.strip().replace(' ', '_').lower()
        if not safe_lang_suffix:
            safe_lang_suffix = "eng"
        unique_sub_id = f"{stremio_id}_{safe_lang_suffix}"

        # URL encode the spaces in the path for Stremio
        raw_url = RAW_BASE_URL + urllib.parse.quote(path)

        if stremio_id not in subtitles_dict:
            subtitles_dict[stremio_id] = []
            
        # --- NEW: Prevent duplicate IDs if both "He" and "he" files exist ---
        existing_ids = [sub["id"] for sub in subtitles_dict[stremio_id]]
        original_unique_id = unique_sub_id
        counter = 2
        
        while unique_sub_id in existing_ids:
            unique_sub_id = f"{original_unique_id}_{counter}"
            counter += 1
        # --------------------------------------------------------------------
        subtitles_dict[stremio_id].append({
            "id": unique_sub_id,
            "url": raw_url,
            "lang": lang_code 
        })

    print(f"[+] Processed {len(subtitles_dict)} episodes with subtitles.")
    os.makedirs(os.path.dirname(OUTPUT_JSON), exist_ok=True)
    
    # --- NEW: Sort the subtitles alphabetically by their 3-letter lang code! ---
    for ep_id in subtitles_dict:
        subtitles_dict[ep_id].sort(key=lambda x: x["lang"])
    # --------------------------------------------------------------------------
    
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(subtitles_dict, f, indent=4, ensure_ascii=False)
    
    print("[+] subtitles.json created successfully!")

if __name__ == "__main__":
    main()
