import os
import re
import json
import time
import random
import requests
import openpyxl
import datetime
from bs4 import BeautifulSoup

SHEET_EXPORT_URL = "https://docs.google.com/spreadsheets/d/1HQRMJgu_zArp-sLnvFMDzOyjdsht87eFLECxMK858lA/export?format=xlsx"
LOCAL_EXCEL_FILE = "one_pace.xlsx"
TRACKER_FILE = "tracker.json"

# --- Load Central Config ---
with open('config.json', 'r', encoding='utf-8') as f:
    CONFIG = json.load(f)
PREFIX_MAP = CONFIG["ARC_MAP"]

def download_excel_file(url, filename, max_retries=3):
    print(f"Downloading latest spreadsheet to {filename}...")
    
    for attempt in range(1, max_retries + 1):
        try:
            response = requests.get(url, stream=True, timeout=20)
            response.raise_for_status()
            
            with open(filename, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk: 
                        f.write(chunk)
                        
            print("Download complete!\n")
            return True
            
        except Exception as e:
            print(f"  [!] Attempt {attempt}/{max_retries} failed: {e}")
            if attempt < max_retries:
                print("  [*] Retrying in 3 seconds...")
                time.sleep(3)
            else:
                print("  [-] All attempts to download the spreadsheet failed.")
                return False

# --- NEW: Upgraded with Retry Logic to respect Nyaa's rate limits ---
def get_torrent_data(nyaa_url, expected_ep_num, max_retries=3):
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    
    for attempt in range(1, max_retries + 1):
        try:
            # Increased timeout slightly to 15 seconds to give Nyaa time to respond
            response = requests.get(nyaa_url, headers=headers, timeout=15)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')
            
            info_hash = None
            torrent_filename = "Unknown Title"

            # 1. Grab the Magnet Hash
            magnet_tag = soup.find('a', href=re.compile(r'^magnet:\?xt=urn:btih:'))
            if magnet_tag:
                match = re.search(r'urn:btih:([a-zA-Z0-9]{40})', magnet_tag['href'])
                if match:
                    info_hash = match.group(1).lower()

            # 2. Grab the Default Filename
            title_tag = soup.find('title')
            if title_tag:
                raw_title = title_tag.text
                torrent_filename = raw_title.replace(" :: Nyaa", "").strip()

            # 3. MAGIC: Dig into the Nyaa File List
            file_list_div = soup.find('div', class_=re.compile('torrent-file-list'))
            if file_list_div:
                ep_padded = str(expected_ep_num).zfill(2)
                lines = file_list_div.get_text(separator='\n').split('\n')
                
                video_files = []
                for line in lines:
                    text = line.strip()
                    if ".mkv" in text or ".mp4" in text:
                        clean_file = re.sub(r'\s*\([^)]*\)$', '', text).strip()
                        video_files.append(clean_file)
                
                if video_files:
                    if len(video_files) == 1:
                        torrent_filename = video_files[0]
                    else:
                        matched = False
                        for vf in video_files:
                            if re.search(rf'\b{ep_padded}\b|\b{expected_ep_num}\b', vf):
                                torrent_filename = vf
                                matched = True
                                break
                        
                        if not matched:
                            ep_index = int(expected_ep_num) - 1
                            if 0 <= ep_index < len(video_files):
                                torrent_filename = video_files[ep_index]
                            else:
                                torrent_filename = video_files[0]

            if info_hash:
                return info_hash, torrent_filename

        # If Nyaa drops the connection, catch the error here!
        except requests.exceptions.RequestException as e:
            print(f"  [!] Nyaa timeout on attempt {attempt}/{max_retries}. Error: {e}")
            if attempt < max_retries:
                # Sleep for 3 to 6 seconds before trying again to let Nyaa cool down
                cooldown = random.uniform(3, 6)
                print(f"  [*] Cooling down for {round(cooldown, 1)} seconds before retrying...")
                time.sleep(cooldown)
            else:
                print(f"  [-] Completely failed to fetch {nyaa_url} after {max_retries} attempts.")
                
    return None, None

def get_expected_filename(ep_name, arc_name):
    prefix = PREFIX_MAP.get(arc_name, arc_name[:2].upper())
    ep_name_str = str(ep_name).strip()
    
    match = re.search(r'(\d+)\s*$', ep_name_str)
    if not match:
        match = re.search(r'\b(\d{1,3})\b', ep_name_str)
        
    if match:
        ep_num = match.group(1)
    else:
        ep_num = "1"
        
    ep_num_int = int(ep_num)
    ep_num = str(ep_num_int)
        
    return f"{prefix}_{ep_num}.json"

def save_tracker(tracker_data):
    with open(TRACKER_FILE, 'w') as f:
        json.dump(tracker_data, f, indent=2)

def main():
    start_time = time.time()
    new_files_count = 0 
    
    output_dir = "stream"
    os.makedirs(output_dir, exist_ok=True)
    
    tracker_data = {}
    if os.path.exists(TRACKER_FILE):
        with open(TRACKER_FILE, 'r') as f:
            tracker_data = json.load(f)
            for key, val in tracker_data.items():
                if isinstance(val, str):
                    tracker_data[key] = [val]
            
    success = download_excel_file(SHEET_EXPORT_URL, LOCAL_EXCEL_FILE)
    if not success:
        return

    print("Loading local spreadsheet...")
    workbook = openpyxl.load_workbook(LOCAL_EXCEL_FILE)
    
    files_to_process = {}
    episode_lengths = {} 
    
    for target_sheet in workbook.sheetnames:
        if target_sheet not in PREFIX_MAP:
            continue
            
        sheet = workbook[target_sheet]
        print(f"\n--- Scanning Arc: {target_sheet} ---")
        
        ep_col_idx, header_row = None, None
        length_col_indices = [] 
        
        for row in range(1, 10):
            for col in range(1, sheet.max_column + 1):
                cell_val = str(sheet.cell(row=row, column=col).value).strip()
                if "One Pace Episode" in cell_val: 
                    ep_col_idx = col
                    header_row = row
                elif "Length" in cell_val: 
                    length_col_indices.append(col)
                    
        if not ep_col_idx:
            print(f"  [!] Could not find the Episode Name column. Skipping.")
            continue

        for row in range(header_row + 1, sheet.max_row + 1):
            ep_name = sheet.cell(row=row, column=ep_col_idx).value
            if not ep_name:
                continue
            
            filename = get_expected_filename(ep_name, target_sheet)
            
            row_lengths = []
            for l_col in length_col_indices:
                val = sheet.cell(row=row, column=l_col).value
                if val:
                    if isinstance(val, datetime.time) or isinstance(val, datetime.datetime):
                        row_lengths.append(val.strftime("%M:%S"))
                    else:
                        row_lengths.append(str(val).strip())

            if filename not in files_to_process:
                files_to_process[filename] = []
                episode_lengths[filename] = {} 
            
            row_urls = []
            for col in range(1, sheet.max_column + 1):
                if col == ep_col_idx: 
                    continue 
                cell = sheet.cell(row=row, column=col)
                target = None
                
                if cell.hyperlink and cell.hyperlink.target:
                    target = cell.hyperlink.target
                elif cell.value and isinstance(cell.value, str) and 'HYPERLINK' in cell.value:
                    match = re.search(r'HYPERLINK\("([^"]+)"', cell.value)
                    if match:
                        target = match.group(1)
                        
                if target and ("nyaa.si" in target or "magnet:" in target):
                    if target not in row_urls:
                        row_urls.append(target)

            for idx, url in enumerate(row_urls):
                if url not in files_to_process[filename]:
                    files_to_process[filename].append(url)
                    assigned_length = row_lengths[idx] if idx < len(row_lengths) else (row_lengths[0] if len(row_lengths) > 0 else "")
                    episode_lengths[filename][url] = assigned_length

    print("\n--- Processing Streams & Saving JSONs ---")
    
    for filename, nyaa_urls in files_to_process.items():
        if not nyaa_urls:
            continue
            
        filepath = os.path.join(output_dir, filename)
        
        if tracker_data.get(filename) == nyaa_urls and os.path.exists(filepath):
            print(f"  [~] Skipped {filename} (Already up-to-date)")
            continue
            
        print(f"  [*] Processing {filename} (Found {len(nyaa_urls)} stream(s)!)")
        
        # --- Extact the raw number from the JSON filename to pass to the scraper! ---
        # e.g., "RO_1.json" -> "1"
        ep_num_raw = filename.split('_')[-1].replace('.json', '')
        
        streams = []
        for url in nyaa_urls:
            # Pass the episode number so it knows exactly which file to grab!
            info_hash, torrent_filename = get_torrent_data(url, ep_num_raw)
            
            if info_hash:
                url_length = episode_lengths.get(filename, {}).get(url, "")
                
                streams.append({
                    "infoHash": info_hash, 
                    "filename": torrent_filename, 
                    "length": url_length
                })
                time.sleep(random.uniform(1, 3)) 
        
        if streams:
            data = {"streams": streams}
            with open(filepath, 'w') as f:
                json.dump(data, f, indent=2)
            
            tracker_data[filename] = nyaa_urls
            save_tracker(tracker_data)
            
            print(f"  [+] Saved {filename}")
            new_files_count += 1
        else:
            print(f"  [-] Failed to get any infoHashes for {filename}")

    end_time = time.time()
    total_time = round(end_time - start_time, 2)
    
    print("\n=========================================")
    print(" ✅ SCRIPT FINISHED ")
    print("=========================================")
    print(f" ⏱️  Time taken: {total_time} seconds")
    print(f" 📥 New files downloaded: {new_files_count}")
    print("=========================================\n")

if __name__ == "__main__":
    main()
