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
try:
    with open('config.json', 'r', encoding='utf-8') as f:
        CONFIG = json.load(f)
except (FileNotFoundError, json.JSONDecodeError) as e:
    print(f"CRITICAL ERROR: Failed to load config.json. {e}")
    exit(1)
PREFIX_MAP = CONFIG["ARC_MAP"]

# Global caches to prevent spamming Nyaa
resolved_batches_cache = {}
nyaa_html_cache = {}

def download_excel_file(url, filename, max_retries=3):
    print(f"Downloading latest spreadsheet to {filename}...")
    for attempt in range(1, max_retries + 1):
        try:
            response = requests.get(url, stream=True, timeout=60)
            response.raise_for_status()
            with open(filename, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk: f.write(chunk)
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

def resolve_nyaa_url(query, max_retries=2):
    rss_url = f"https://nyaa.si/?page=rss&q={query}&c=0_0&f=0"
    for attempt in range(1, max_retries + 1):
        try:
            response = requests.get(rss_url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=15)
            response.raise_for_status()
            match = re.search(r'<guid isPermaLink="true">(https://nyaa\.si/view/\d+)</guid>', response.text)
            if match:
                return match.group(1)
            break 
        except requests.exceptions.RequestException:
            if attempt < max_retries: time.sleep(random.uniform(1, 3))

    html_url = f"https://nyaa.si/?q={query}&f=0&c=0_0"
    for attempt in range(1, max_retries + 1):
        try:
            response = requests.get(html_url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=15)
            response.raise_for_status()
            match = re.search(r'href="(/view/\d+)"', response.text)
            if match:
                return f"https://nyaa.si{match.group(1)}"
            return None
        except requests.exceptions.RequestException:
            if attempt < max_retries: time.sleep(random.uniform(1, 3))
    return None

def resolve_nyaa_batch(arc_name, max_retries=3):
    aliases = {
        "Alabasta": "Arabasta",
        "Whisky Peak": "Whiskey Peak",
        "Return to Sabaody": "Sabaody",
        "The Adventures of Buggys Crew": "Buggy",
        "The Trials of Koby-Meppo": "Koby-Meppo"
    }
    search_name = aliases.get(arc_name, arc_name)
    clean_name = search_name.replace("The Adventures of ", "").replace("The Trials of ", "").strip()
    
    # Search globally without forcing 1080p so we catch the older 720p/480p arcs!
    query = f"One Pace {clean_name}".replace("'", "").replace(' ', '+')
    html_url = f"https://nyaa.si/?f=0&c=0_0&q={query}"
    
    for attempt in range(1, max_retries + 1):
        try:
            response = requests.get(html_url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=15)
            response.raise_for_status()
            matches = re.findall(r'<a href="(/view/\d+)" title="([^"]+)">', response.text)
            
            best_link = None
            best_score = -1
            
            for link, title in matches:
                if "One Pace" not in title or clean_name not in title or link.endswith("#comments"):
                    continue
                
                score = 0
                if "1080" in title: score += 30
                elif "720" in title: score += 20
                elif "480" in title: score += 10
                
                # --- SMART BATCH DETECTION ---
                is_batch = False
                if "Batch" in title or "batch" in title.lower():
                    is_batch = True
                elif re.search(rf'{clean_name}\s+\d{{1,3}}\s*[-~]\s*\d{{1,3}}', title, re.IGNORECASE):
                    is_batch = True # Catches "Arabasta 01-21"
                elif not re.search(rf'{clean_name}\s+\d{{1,3}}\b', title, re.IGNORECASE):
                    is_batch = True # Catches "Arabasta [1080p]" (No isolated episode number)
                    
                if is_batch:
                    score += 100 
                    
                if score > best_score:
                    best_score = score
                    best_link = f"https://nyaa.si{link}"

            # CRITICAL FIX: Only return the URL if we are CERTAIN it is a batch!
            if best_link and best_score >= 100:
                print(f"  [✅] Found Best Batch URL (Score {best_score}): {best_link}")
                return best_link
                
            print(f"  [-] No batch found for {arc_name}")
            return None
        except requests.exceptions.RequestException:
            if attempt < max_retries: time.sleep(random.uniform(2, 4))
    return None

def resolve_single_episode(arc_name, ep_num, max_retries=2):
    aliases = {
        "Alabasta": "Arabasta",
        "Whisky Peak": "Whiskey Peak",
        "Return to Sabaody": "Sabaody",
        "The Adventures of Buggys Crew": "Buggy",
        "The Trials of Koby-Meppo": "Koby-Meppo"
    }
    search_name = aliases.get(arc_name, arc_name)
    clean_name = search_name.replace("The Adventures of ", "").replace("The Trials of ", "").strip()
    
    ep_padded = str(ep_num).zfill(2)
    query = f"One Pace {clean_name} {ep_padded}".replace("'", "").replace(' ', '+')
    html_url = f"https://nyaa.si/?f=0&c=0_0&q={query}"
    
    for attempt in range(1, max_retries + 1):
        try:
            response = requests.get(html_url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=15)
            response.raise_for_status()
            matches = re.findall(r'<a href="(/view/\d+)" title="([^"]+)">', response.text)
            
            best_link = None
            best_score = -1
            
            for link, title in matches:
                if "One Pace" not in title or clean_name not in title or link.endswith("#comments"):
                    continue
                
                # Must contain the explicit episode number!
                if not re.search(rf'\b{ep_padded}\b|\b{ep_num}\b', title):
                    continue
                    
                score = 0
                if "1080" in title: score += 30
                elif "720" in title: score += 20
                elif "480" in title: score += 10
                
                if score > best_score:
                    best_score = score
                    best_link = f"https://nyaa.si{link}"

            if best_link:
                print(f"  [✅] Found Single Episode URL: {best_link}")
                return best_link
            return None
        except requests.exceptions.RequestException:
            if attempt < max_retries: time.sleep(random.uniform(1, 3))
    return None

def get_torrent_data(nyaa_url, expected_ep_num, expected_crc=None, max_retries=3):
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    html_content = None
    
    if nyaa_url in nyaa_html_cache:
        html_content = nyaa_html_cache[nyaa_url]
    else:
        for attempt in range(1, max_retries + 1):
            try:
                response = requests.get(nyaa_url, headers=headers, timeout=15)
                response.raise_for_status()
                html_content = response.text
                nyaa_html_cache[nyaa_url] = html_content
                break
            except requests.exceptions.RequestException as e:
                print(f"  [!] Nyaa timeout on attempt {attempt}/{max_retries}. Error: {e}")
                if attempt < max_retries:
                    time.sleep(random.uniform(3, 6))
                else:
                    return None, None

    if not html_content:
        return None, None

    soup = BeautifulSoup(html_content, 'html.parser')
    info_hash = None
    torrent_filename = "Unknown Title"

    magnet_tag = soup.find('a', href=re.compile(r'^magnet:\?xt=urn:btih:'))
    if magnet_tag:
        match = re.search(r'urn:btih:([a-zA-Z0-9]{40})', magnet_tag['href'])
        if match:
            info_hash = match.group(1).lower()

    title_tag = soup.find('title')
    if title_tag:
        raw_title = title_tag.text
        torrent_filename = raw_title.replace(" :: Nyaa", "").strip()

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
            video_files.sort()
            matched = False
            
            # --- STRICT CRC MATCHING ---
            if expected_crc and expected_crc not in ["Unknown", "00000000"]:
                for vf in video_files:
                    if expected_crc.upper() in vf.upper() or expected_crc.lower() in vf.lower():
                        torrent_filename = vf
                        matched = True
                        break
                        
                if not matched:
                    # REJECT: We expected a precise CRC and this torrent doesn't have it!
                    return None, None
            
            # --- FALLBACK: Episode Number Matching (Only if no CRC provided) ---
            if not matched:
                if len(video_files) == 1:
                    torrent_filename = video_files[0]
                    matched = True
                else:
                    for vf in video_files:
                        if re.search(rf'\b{ep_padded}\b|\b{expected_ep_num}\b', vf):
                            torrent_filename = vf
                            matched = True
                            break
                            
            # --- FINAL FALLBACK: Index ---
            if not matched:
                ep_index = int(expected_ep_num) - 1
                if 0 <= ep_index < len(video_files):
                    torrent_filename = video_files[ep_index]
                else:
                    # CRITICAL FIX: If out of bounds, do NOT guess the first file! Reject it!
                    return None, None

    if info_hash:
        return info_hash, torrent_filename
    return None, None

def get_expected_filename(ep_name, arc_name):
    prefix = PREFIX_MAP.get(arc_name, arc_name[:2].upper())
    ep_name_str = str(ep_name).strip()
    match = re.search(r'(\d+)\s*$', ep_name_str)
    if not match: match = re.search(r'\b(\d{1,3})\b', ep_name_str)
    ep_num = match.group(1) if match else "1"
    return f"{prefix}_{str(int(ep_num))}.json"

def save_tracker(tracker_data):
    with open(TRACKER_FILE, 'w') as f:
        json.dump(tracker_data, f, indent=2)

def resolve_from_website(arc_name, ep_num_raw, max_retries=2):   
    url = "https://onepace.net/en/releases"
    for attempt in range(1, max_retries + 1):
        try:
            response = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=15)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Normalize names just like in the other resolve functions
            aliases = {
                "Alabasta": "Arabasta",
                "Whisky Peak": "Whiskey Peak",
                "Return to Sabaody": "Sabaody",
                "The Adventures of Buggys Crew": "Buggy",
                "The Trials of Koby-Meppo": "Koby-Meppo"
            }
            search_name = aliases.get(arc_name, arc_name)
            clean_name = search_name.replace("The Adventures of ", "").replace("The Trials of ", "").strip().lower()
            ep_padded = str(ep_num_raw).zfill(2)

            # All episodes are grouped in <li> tags that contain an ID
            episodes = soup.find_all('li', id=True) 
            
            for ep in episodes:
                title_tag = ep.find('h3')
                if not title_tag: continue
                
                title_text = title_tag.get_text(separator=" ", strip=True).lower()
                
                # [FIXED] Allows matching either zero-padded (e.g., "05") or raw (e.g., "5")
                if clean_name not in title_text:
                    continue
                if not re.search(rf'\b{ep_padded}\b|\b{ep_num_raw}\b', title_text):
                    continue
                
                # Check if it was released within the last 7 days (1 week limit)
                time_tag = ep.find('time')
                if not time_tag or not time_tag.has_attr('datetime'):
                    continue
                
                try:
                    # Example format: 2026-04-22
                    release_date = datetime.datetime.strptime(time_tag['datetime'][:10], "%Y-%m-%d").date()
                    today = datetime.datetime.now().date()
                    days_diff = (today - release_date).days
                    
                    if days_diff > 7 or days_diff < 0:
                        continue # Older than 1 week, ignore it
                except Exception:
                    continue
                
                # --- CONDITION 1: Look for the Magnet Link ---
                magnet_tag = ep.find('a', href=re.compile(r'^magnet:\?xt='))
                if magnet_tag:
                    match = re.search(r'urn:btih:([a-zA-Z0-9]{40})', magnet_tag['href'], re.IGNORECASE)
                    if match:
                        info_hash = match.group(1).lower()
                        print(f"  [🌐] Fresh episode on Website (Magnet): {info_hash}")
                        # Resolve it into a Nyaa View Link so get_torrent_data handles it smoothly
                        view_url = resolve_nyaa_url(info_hash)
                        if view_url: return view_url
                
                # --- CONDITION 2: Look for the Torrent Link ---
                torrent_tag = ep.find('a', href=re.compile(r'nyaa\.si/download/\d+\.torrent'))
                if torrent_tag:
                    match = re.search(r'/download/(\d+)\.torrent', torrent_tag['href'])
                    if match:
                        nyaa_id = match.group(1)
                        print(f"  [🌐] Fresh episode on Website (Torrent): {nyaa_id}")
                        # Create the Nyaa view link 
                        return f"https://nyaa.si/view/{nyaa_id}"
                        
            return None
        except requests.exceptions.RequestException:
            if attempt < max_retries: time.sleep(random.uniform(1, 3))
    return None

def main():
    start_time = time.time()
    new_files_count = 0 
    updated_files_list = [] 
    
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
    if not success: return

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
            if not ep_name: continue
            
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
                if col == ep_col_idx: continue 
                cell = sheet.cell(row=row, column=col)
                target = None
                
                if cell.hyperlink and cell.hyperlink.target:
                    target = cell.hyperlink.target
                elif cell.value and isinstance(cell.value, str) and 'HYPERLINK' in cell.value:
                    match = re.search(r'HYPERLINK\("([^"]+)"', cell.value)
                    if match: target = match.group(1)
                elif cell.value and isinstance(cell.value, str):
                    val = cell.value.strip()
                    if re.match(r'^[A-Fa-f0-9]{8}$', val):
                        target = f"BATCH_SEARCH:{target_sheet}:{val}"
                    elif val.isdigit() and 6 <= len(val) <= 8:
                        target = f"https://nyaa.si/view/{val}"
                    # --- NEW: Catch raw 40-character infohashes just in case! ---
                    elif re.match(r'^[A-Fa-f0-9]{40}$', val):
                        target = f"https://nyaa.si/?q={val}"
                        
                if target and ("nyaa.si" in target or "magnet:" in target or target.startswith("BATCH_SEARCH:")):
                    if target not in row_urls:
                        row_urls.append(target)

            # --- SMART FAILSAFE: Prevent guessing unreleased episodes ---
            if not row_urls:
                # Unreleased episodes have no video length. 
                # If it HAS a length but no link, it's a typo (like SK_14), so we hunt for it!
                has_valid_length = any(l and str(l).strip() and "TBA" not in str(l).upper() for l in row_lengths)
                if has_valid_length:
                    row_urls.append(f"BATCH_SEARCH:{target_sheet}:Unknown")

            for idx, url in enumerate(row_urls):
                if url not in files_to_process[filename]:
                    files_to_process[filename].append(url)
                    assigned_length = row_lengths[idx] if idx < len(row_lengths) else ""
                    episode_lengths[filename][url] = assigned_length

    print("\n--- Processing Streams & Saving JSONs ---")
    
    for filename, nyaa_urls in files_to_process.items():
        if not nyaa_urls: continue
            
        filepath = os.path.join(output_dir, filename)
        
        if tracker_data.get(filename) == nyaa_urls and os.path.exists(filepath):
            print(f"  [~] Skipped {filename} (Already up-to-date)")
            continue
            
        print(f"  [*] Processing {filename} (Found {len(nyaa_urls)} stream(s)!)")
        ep_num_raw = filename.split('_')[-1].replace('.json', '')
        
        streams = []
        for url in nyaa_urls:
            info_hash, torrent_filename = None, None
            actual_url = url
            expected_crc = None
            
            # --- PHASE 1: Handle raw InfoHash queries from the spreadsheet ---
            if "nyaa.si/?q=" in actual_url:
                match = re.search(r'\?q=([a-fA-F0-9]{40})', actual_url)
                if match:
                    hash_val = match.group(1)
                    print(f"  [🔎] Resolving Spreadsheet InfoHash: {hash_val}...")
                    resolved = resolve_nyaa_url(hash_val)
                    if resolved:
                        actual_url = resolved
                    else:
                        actual_url = f"BATCH_SEARCH:{target_sheet}:Unknown" 
            
            # --- PHASE 2: Test Direct URLs (If it's not a batch search) ---
            if not actual_url.startswith("BATCH_SEARCH:"):
                info_hash, torrent_filename = get_torrent_data(actual_url, ep_num_raw, expected_crc=None)
                if not info_hash:
                    print(f"  [⏳] Direct link failed/deleted. Triggering Fallback Chain...")
                    actual_url = f"BATCH_SEARCH:{target_sheet}:Unknown"
            
            # --- PHASE 3: The Ultimate Fallback Chain ---
            if actual_url.startswith("BATCH_SEARCH:"):
                parts = actual_url.split(":")
                arc_name = parts[1]
                expected_crc = parts[2]
                
                # TIER 0: OFFICIAL WEBSITE BACKUP (Fresh Releases <= 7 Days)
                website_url = resolve_from_website(arc_name, ep_num_raw)
                if website_url:
                    info_hash, torrent_filename = get_torrent_data(website_url, ep_num_raw, expected_crc)
                    if info_hash: time.sleep(random.uniform(1, 2))
                
                # TIER 1: ALWAYS TRY THE OFFICIAL ARC BATCH FIRST
                if not info_hash:
                    batch_url = resolved_batches_cache.get(arc_name)
                    if not batch_url:
                        batch_url = resolve_nyaa_batch(arc_name)
                        if batch_url:
                            resolved_batches_cache[arc_name] = batch_url
                            time.sleep(random.uniform(1, 2))
                    
                    if batch_url:
                        info_hash, torrent_filename = get_torrent_data(batch_url, ep_num_raw, expected_crc)
                
                # TIER 2: TRY PRECISE CRC SEARCH (If missing from batch)
                if not info_hash and expected_crc not in ["Unknown", "00000000"]:
                    print(f"  [🔎] Not in batch. Resolving CRC directly: {expected_crc}...")
                    crc_url = resolve_nyaa_url(expected_crc)
                    if crc_url:
                        info_hash, torrent_filename = get_torrent_data(crc_url, ep_num_raw, expected_crc)
                        if info_hash: time.sleep(random.uniform(1, 2))
                            
                # TIER 3: SINGLE EPISODE SEARCH (Final Fallback)
                if not info_hash:
                    print(f"  [⏳] CRC missed. Hunting for Single Episode: {arc_name} {ep_num_raw}...")
                    sing_url = resolve_single_episode(arc_name, ep_num_raw)
                    if sing_url:
                        info_hash, torrent_filename = get_torrent_data(sing_url, ep_num_raw, None)
                        if info_hash: time.sleep(random.uniform(1, 2))
                
                if not info_hash:
                    print(f"  [-] ❌ Could not resolve torrent for CRC, Episode, or Batch: {arc_name}")
                    continue
                
            if info_hash:
                url_length = episode_lengths.get(filename, {}).get(url, "")
                streams.append({
                    "infoHash": info_hash, 
                    "filename": torrent_filename, 
                    "length": url_length
                })
                # Only sleep if we made a successful direct network hit
                if actual_url not in nyaa_html_cache:
                    time.sleep(random.uniform(0.5, 1.5))
        
        if streams:
            data = {"streams": streams}
            with open(filepath, 'w') as f:
                json.dump(data, f, indent=2)
            
            tracker_data[filename] = nyaa_urls
            save_tracker(tracker_data)
            
            print(f"  [+] Saved {filename}")
            new_files_count += 1
            updated_files_list.append(filename) 
            
        else:
            print(f"  [-] Failed to get any infoHashes for {filename}")

    if updated_files_list:
        with open("stream/st_purge.txt", "w") as f:
            for fname in updated_files_list:
                f.write(fname + "\n")

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