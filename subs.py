import urllib.request
import json
import re
import urllib.parse
import os
import time

# --- Configuration ---
REPO_API_URL = "https://api.github.com/repos/one-pace/one-pace-public-subtitles/git/trees/main?recursive=1"

#RAW_ASS_BASE_URL = "https://cdn.jsdelivr.net/gh/one-pace/one-pace-public-subtitles@main/"
RAW_ASS_BASE_URL = "https://raw.githubusercontent.com/one-pace/one-pace-public-subtitles/main/"

CDN_SRT_BASE_URL = "https://cdn.jsdelivr.net/gh/6ip/onepace-streams@main/meta/subs/"

OUTPUT_JSON = "meta/subtitles.json"
OUTPUT_SUBS_DIR = "meta/subs"
HASHES_FILE = "hashes.json"

# --- Load Central Config ---
with open('config.json', 'r', encoding='utf-8') as f:
    CONFIG = json.load(f)
ARC_MAP = CONFIG["ARC_MAP"]

LANG_MAP = {
    "ar": "ara", "cs": "cze", "cz": "cze", "de": "ger", "en": "eng", 
    "en cc": "eng", "en-cc": "eng", "en_cc": "eng",
    "es": "spa", "fi": "fin", "fr": "fre", "he": "heb", "it": "ita", "pl": "pol",
    "pt": "por", "ptbr": "por", "pt-br": "por", "pt_br": "por",
    "ru": "rus", "tr": "tur", "id": "ind", "nl": "dut", "vi": "vie",
    "ja": "jpn", "typesetting": "eng"
}

# --- Conversion Tools ---
_TIME_PATTERN = re.compile(r"(\d+):(\d{2}):(\d{2})\.(\d{2})")
_ASS_TAG_PATTERN = re.compile(r"\{[^}]*\}")
_SKIP_STYLES = ["sign", "song", "op ", "ed ", "karaoke", "chapter", "credit", "eyecatch", "next ep", "preview"]

def convert_time(t: str) -> str:
    match = _TIME_PATTERN.match(t)
    if match:
        h, m, s, cs = match.groups()
        return f"{int(h):02d}:{m}:{s},{cs}0"
    return t

def ass_to_srt(ass_content: str) -> str:
    lines = ass_content.split("\n")
    events_section = False
    format_line = None
    dialogues = []

    for line in lines:
        line = line.strip()
        if line == "[Events]":
            events_section = True
            continue
        if line.startswith("[") and line.endswith("]") and events_section:
            events_section = False
            continue
        if events_section and line.startswith("Format:"):
            format_line = [f.strip() for f in line[7:].split(",")]
            continue
        if events_section and line.startswith("Dialogue:"):
            if format_line:
                parts = line[10:].split(",", len(format_line) - 1)
                if len(parts) == len(format_line):
                    entry = dict(zip(format_line, parts))
                    style = entry.get("Style", "").lower()
                    if any(s in style for s in _SKIP_STYLES):
                        continue
                    dialogues.append(entry)

    # --- NEW: Helper to calculate milliseconds for accurate sorting ---
    def time_to_ms(t_str):
        match = _TIME_PATTERN.match(t_str)
        if match:
            h, m, s, cs = match.groups()
            return int(h) * 3600000 + int(m) * 60000 + int(s) * 1000 + int(cs) * 10
        return 0

    # --- NEW: Sort the dialogue chronologically so the SRT doesn't break ---
    dialogues.sort(key=lambda x: time_to_ms(x.get("Start", "0:00:00.00")))

    srt_lines = [
        "1",
        "00:00:01,000 --> 00:00:04,000",
        r'{\an8}<b><font color="#9CD5FF">One Pace Premium</font></b>',
        r'Keep the project alive: <font color="#a8c7fa">ko-fi.com/not6ip</font>',
        ""
    ]
    
    # --- NEW: Group texts by exact timestamp to merge fragmented signs ---
    grouped_dialogues = {}

    for d in dialogues:
        start = d.get("Start", "0:00:00.00")
        end = d.get("End", "0:00:00.00")
        text = d.get("Text", "")

        # Skip lines that are vector drawings (e.g., {\p1}m 848 508 l...)
        if r"\p1" in text or r"\p2" in text or r"\p4" in text:
            continue

        # Scrub out Aegisub typesetter "invisible text" background hacks
        # This deletes the invisibility tag (even if buried in other tags) AND the garbage text that follows it
        text = re.sub(r'\{[^}]*\\[1-4]?(alpha|a)&H[Ff]{2}&[^}]*\}[^{]*', '', text)

        text = _ASS_TAG_PATTERN.sub("", text)
        text = text.replace("\\N", "\n").replace("\\n", "\n").strip()
        
        # Strip invisible bidirectional formatting characters (RLE, LRE, etc.)
        text = re.sub(r'[\u200e\u200f\u202a\u202b\u202c\u202d\u202e]', '', text).strip()

        if not text or "mpv.io" in text.lower():
            continue

        # Group by timestamp to catch split lines
        time_key = (start, end)
        if time_key not in grouped_dialogues:
            grouped_dialogues[time_key] = []
        
        # Deduplicate identical layers
        if text not in grouped_dialogues[time_key]:
            grouped_dialogues[time_key].append(text)

    # --- NEW: Write the grouped dialogues to the SRT format ---
    counter = 2
    for (start, end), texts in grouped_dialogues.items():
        srt_lines.append(str(counter))
        srt_lines.append(f"{convert_time(start)} --> {convert_time(end)}")
        srt_lines.append("\n".join(texts))  # Join multiple lines sharing the same time!
        srt_lines.append("")
        counter += 1

    return "\n".join(srt_lines)

def clean_string(s):
    return re.sub(r'[\d\s\-]', '', s).lower()

# --- Main Logic ---
def main():
    print("[?] Fetching subtitle repository tree from GitHub...")
    req = urllib.request.Request(REPO_API_URL, headers={'User-Agent': 'Mozilla/5.0'})
    
    try:
        with urllib.request.urlopen(req) as response:
            tree_data = json.loads(response.read().decode())
    except Exception as e:
        print(f"[-] Failed to fetch from GitHub API: {e}")
        return

    local_hashes = {}
    if os.path.exists(HASHES_FILE):
        with open(HASHES_FILE, "r", encoding="utf-8") as f:
            try:
                local_hashes = json.load(f)
            except json.JSONDecodeError:
                print("[-] Could not read hashes.json, starting fresh.")

    os.makedirs(OUTPUT_SUBS_DIR, exist_ok=True)
    subtitles_dict = {}
    
    ass_files = [item for item in tree_data.get("tree", []) if item.get("path", "").endswith(".ass")]
    print(f"[+] Found {len(ass_files)} .ass files to process. Starting conversion...")

    for i, item in enumerate(ass_files):
        path = item.get("path", "")
        # --- NEW: Explicitly ignore the Release folder ---
        if "Release/" in path or "Final Subs/" in path:
            continue
        file_sha = item.get("sha", "") 
        
        parts = path.split("/")
        if len(parts) < 3: continue
        
        filename = parts[-1]        
        ep_folder = parts[-2]       
        arc_folder = parts[-3]      

        arc_key = clean_string(arc_folder)
        prefix = ARC_MAP.get(arc_key)

        # --- NEW: Fallback for Cover Stories & Specials ---
        # If the folder name doesn't map to a prefix (like "00 Cover Stories"), 
        # we check if the actual filename contains the arc name (like "Buggys Crew")
        if not prefix:
            clean_fname = clean_string(filename)
            for key, val in ARC_MAP.items():
                clean_k = clean_string(key)
                if clean_k and clean_k in clean_fname:
                    prefix = val
                    break

        if not prefix: continue

        try:
            ep_num = int(ep_folder)
        except ValueError:
            continue

        # --- NEW: Force episode ID to 1 for Cover Stories & Specials ---
        if "Cover Stories" in arc_folder or prefix in ["BUGGYS_CREW", "COVER_KOBYMEPPO", "COVER_SHSS"]:
            stremio_id = f"{prefix}_1"
        else:
            stremio_id = f"{prefix}_{ep_num}"

        name_without_ext = filename.rsplit('.', 1)[0]
        ep_str = str(ep_num).zfill(2)
        idx = name_without_ext.rfind(ep_str)
        
        # --- NEW: Smarter Language Extraction ---
        if idx != -1:
            raw_lang_str = name_without_ext[idx + len(ep_str):].strip()
        else:
            # If no episode number in filename (like the Specials), 
            # we scan the filename from right-to-left to extract language tags/modifiers
            words = name_without_ext.split()
            lang_parts = []
            for w in reversed(words):
                w_lower = w.lower()
                # Check if word is a known language or modifier
                if w_lower in LANG_MAP or w_lower in ['alternate', 'dub', 'cc', 'typesetting']:
                    lang_parts.insert(0, w_lower)
                else:
                    break # Stop when we hit normal title words like "Crew"
            raw_lang_str = " ".join(lang_parts)

        lang_code = "eng" 
        for word in raw_lang_str.lower().split():
            if word in LANG_MAP:
                lang_code = LANG_MAP[word]
                break 
                
        safe_lang_suffix = raw_lang_str.strip().replace(' ', '_').lower() or "eng"
        unique_sub_id = f"{stremio_id}_{safe_lang_suffix}"

        if stremio_id not in subtitles_dict:
            subtitles_dict[stremio_id] = []
            
        existing_ids = [sub["id"] for sub in subtitles_dict[stremio_id]]
        original_unique_id = unique_sub_id
        counter = 2
        while unique_sub_id in existing_ids:
            unique_sub_id = f"{original_unique_id}_{counter}"
            counter += 1

        srt_filename = f"{unique_sub_id}.srt"
        
        # --- NEW: Organize into Arc/Episode folders ---
        nested_dir = os.path.join(OUTPUT_SUBS_DIR, arc_folder, ep_folder)
        os.makedirs(nested_dir, exist_ok=True) # Creates the nested folders if they don't exist
        
        local_srt_path = os.path.join(nested_dir, srt_filename)
        
        # --- NEW: Update the CDN URL to match ---
        # We use safe='/' so the slashes stay as '/' instead of becoming '%2F'
        rel_path = f"{arc_folder}/{ep_folder}/{srt_filename}"
        cdn_url = CDN_SRT_BASE_URL + urllib.parse.quote(rel_path, safe='/')

        # --- NEW LOGIC: Crash-Proof Hash Checking ---
        file_exists = os.path.exists(local_srt_path)
        
        # If the file exists but we don't have its hash (because of Ctrl+C), just update the hash and skip download!
        if file_exists and path not in local_hashes:
            local_hashes[path] = file_sha
            needs_download = False
        else:
            # Download if missing, OR if we have a saved hash and it doesn't match GitHub
            needs_download = not file_exists or (path in local_hashes and local_hashes[path] != file_sha)

        if needs_download:
            download_url = RAW_ASS_BASE_URL + urllib.parse.quote(path)
            
            success = False
            for attempt in range(5):
                try:
                    dl_req = urllib.request.Request(download_url, headers={'User-Agent': 'Mozilla/5.0'})
                    with urllib.request.urlopen(dl_req) as dl_resp:
                        ass_text = dl_resp.read().decode('utf-8-sig', errors='ignore')
                        
                    srt_text = ass_to_srt(ass_text)
                    
                    with open(local_srt_path, "w", encoding="utf-8") as f:
                        f.write(srt_text)
                        
                    local_hashes[path] = file_sha
                        
                    success = True
                    time.sleep(0.5) # Increased to 0.5s to prevent the 403 error from happening again!
                    break 
                    
                except urllib.error.HTTPError as e:
                    if e.code in [429, 403]:
                        wait_time = 3 ** attempt # Wait longer if we get blocked!
                        print(f"    [!] Blocked by server! Cooling down for {wait_time} seconds...")
                        time.sleep(wait_time)
                    else:
                        print(f"[-] HTTP Error {e.code} on {filename}")
                        break
                except Exception as e:
                    print(f"[-] Error converting {filename}: {e}")
                    break
            
            if not success:
                print(f"[-] Giving up on {filename} after 5 retries.")
                continue

        # Add to JSON dict
        subtitles_dict[stremio_id].append({
            "id": unique_sub_id,
            "url": cdn_url,
            "lang": lang_code 
        })

        # --- NEW LOGIC: Autosave hashes every 50 files ---
        if i > 0 and i % 50 == 0:
            print(f"    -> Processing {i}/{len(ass_files)} files... (Autosaving hashes)")
            with open(HASHES_FILE, "w", encoding="utf-8") as f:
                json.dump(local_hashes, f, indent=4)

    # Sort JSON
    for ep_id in subtitles_dict:
        subtitles_dict[ep_id].sort(key=lambda x: x["lang"])
        
    os.makedirs(os.path.dirname(OUTPUT_JSON), exist_ok=True)
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(subtitles_dict, f, indent=4, ensure_ascii=False)
        
    with open(HASHES_FILE, "w", encoding="utf-8") as f:
        json.dump(local_hashes, f, indent=4)
    
    print(f"\n[+] Processed {len(subtitles_dict)} episodes.")
    print("[+] Subtitle files saved to meta/subs/")
    print("[+] hashes.json finalized!")
    print("[+] subtitles.json created successfully!")

if __name__ == "__main__":
    main()
