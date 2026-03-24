import json
import csv
import urllib.request
from io import StringIO

# --- Configuration ---
OUTPUT_JSON = 'meta/pp_onepace.json'
BASE_META_URL = "https://fedew04.github.io/OnePaceStremio/meta/series/pp_onepace.json"

SHEET_ID = "1M0Aa2p5x7NioaH9-u8FyHq6rH3t5s6Sccs8GoC6pHAM"
CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid=0"

# Exact matching dictionary based on your raw CSV data
ARC_PREFIXES = {
    "romancedawn": "RO",
    "orangetown": "OR",
    "syrupvillage": "SY",
    "gaimon": "GA",
    "baratie": "BA",
    "arlongpark": "AR",
    "theadventuresofbuggyscrew": "BUGGYS_CREW",
    "loguetown": "LO",
    "reversemountain": "RM",
    "whiskypeak": "WH", # Fixed spelling!
    "whiskeypeak": "WH", # Fallback just in case
    "thetrialsofkobymeppo": "COVER_KOBYMEPPO",
    "littlegarden": "LI",
    "drumisland": "DI",
    "arabasta": "AL",
    "alabasta": "AL", # Fallback
    "jaya": "JA",
    "skypiea": "SK",
    "longringlongland": "LR",
    "waterseven": "WS",
    "enieslobby": "EN",
    "postenieslobby": "PEN",
    "thrillerbark": "TB",
    "sabaodyarchipelago": "SAB",
    "amazonlily": "AM",
    "impeldown": "IM",
    "ifyoucouldgoanywheretheadventuresofthestrawhats": "COVER_SHSS",
    "marineford": "MA",
    "postwar": "PW",
    "returntosabaody": "RTS",
    "fishmanisland": "FI",
    "punkhazard": "PH",
    "dressrosa": "DR",
    "zou": "ZO",
    "wholecakeisland": "WC",
    "reverie": "REV",
    "wano": "WA",
    "egghead": "EH"
    #"elbaf": "EL" # <--- You would just add the new one here # For new seasons (3 changes)!
}

def clean_string(s):
    """Removes spaces, dashes, apostrophes, and periods for exact matching."""
    return str(s).lower().replace(" ", "").replace("-", "").replace("'", "").replace(".", "")

def main():
    print("1. Fetching original JSON from fedew04...")
    try:
        req = urllib.request.urlopen(BASE_META_URL)
        data = json.loads(req.read().decode('utf-8'))
    except Exception as e:
        print(f"Failed to download base JSON: {e}")
        return

    print("2. Fetching descriptions from Google Sheets...")
    try:
        req = urllib.request.urlopen(CSV_URL)
        csv_content = req.read().decode('utf-8')
    except Exception as e:
        print(f"Failed to download spreadsheet: {e}")
        return

    # Build description dictionary
    descriptions_map = {}
    reader = csv.DictReader(StringIO(csv_content))
    
    for row in reader:
        arc_title = row.get("arc_title", "").strip()
        arc_part = row.get("arc_part", "").strip()
        desc = row.get("description_en", "").strip()

        if not arc_title or not arc_part or not desc:
            continue

        clean_arc = clean_string(arc_title)
        
        # EXACT MATCH - This ignores Specials automatically and prevents overlap
        if clean_arc in ARC_PREFIXES:
            matched_prefix = ARC_PREFIXES[clean_arc]
            
            try:
                # Convert to int to strip leading zeros, then back to string
                arc_part_str = str(int(arc_part))
            except ValueError:
                # Failsafe if the number isn't a standard integer
                arc_part_str = arc_part 

            video_id = f"{matched_prefix}_{arc_part_str}"
            descriptions_map[video_id] = desc

    print("3. Applying server.js transformations...")
    meta = data.get("meta", {})
    
    # Static Data Injection
    meta["poster"] = "https://cdn.jsdelivr.net/gh/6ip/onepace-assets-prm@main/public/poster.jpg"
    meta["background"] = "https://image.tmdb.org/t/p/original/iN5LKyvyWUWwqbjaQfKFXoo8mch.jpg"
    meta["logo"] = "https://cdn.jsdelivr.net/gh/6ip/onepace-assets-prm@main/public/logo.png"
    meta["description"] = "Experience One Piece without the filler. This manga-accurate cut removes padded scenes, saving you hundreds of hours while staying true to Oda's original vision."
    meta["releaseInfo"] = "1999-"
    meta["year"] = "1999-"
    meta["imdbRating"] = "9.0"
    meta["country"] = "Japan"
    meta["released"] = "1999-10-20T00:00:00.000Z"
    meta["genres"] = ["Animation", "Action", "Adventure"]
    meta["cast"] = ["Mayumi Tanaka", "Akemi Okamura", "Tony Beck"]
    
    # Links Array
    meta["links"] = [
        {"name": "9.0", "category": "imdb", "url": "https://imdb.com/title/tt0388629"},
        {"name": "Animation", "category": "Genres", "url": "stremio:///discover/https%3A%2F%2Fv3-cinemeta.strem.io%2Fmanifest.json/series/top?genre=Animation"},
        {"name": "Action", "category": "Genres", "url": "stremio:///discover/https%3A%2F%2Fv3-cinemeta.strem.io%2Fmanifest.json/series/top?genre=Action"},
        {"name": "Adventure", "category": "Genres", "url": "stremio:///discover/https%3A%2F%2Fv3-cinemeta.strem.io%2Fmanifest.json/series/top?genre=Adventure"},
        {"name": "Mayumi Tanaka", "category": "Cast", "url": "stremio:///search?search=Mayumi%20Tanaka"},
        {"name": "Akemi Okamura", "category": "Cast", "url": "stremio:///search?search=Akemi%20Okamura"},
        {"name": "Tony Beck", "category": "Cast", "url": "stremio:///search?search=Tony%20Beck"}
    ]

    # Generate Seasons Array (1 to 33)
    # For new seasons: # Change range(1, 34) to range(1, 35) to include season 34
    seasons_array = []
    for i in range(1, 34):
        s_padded = str(i).zfill(2)
        seasons_array.append({
            "season": i,
            "poster": f"https://cdn.jsdelivr.net/gh/6ip/onepace-assets-prm@main/public/poster-s/poster-s{s_padded}.jpg"
        })
    meta["seasons"] = seasons_array

    # Process Videos (Thumbnails + Descriptions)
    desc_count = 0
    for video in meta.get("videos", []):
        vid_id = video.get("id")
        season_num = video.get("season")

        # Set Thumbnail
        # For new seasons: # Change 33 to 34
        if season_num and 1 <= season_num <= 33:
            s_padded = str(season_num).zfill(2)
            video["thumbnail"] = f"https://images.weserv.nl/?url=cdn.jsdelivr.net/gh/6ip/onepace-assets-prm@main/public/poster-s/poster-s{s_padded}.jpg&w=1280&h=720&fit=cover&a=center"
        else:
            video["thumbnail"] = "https://image.tmdb.org/t/p/w500/iN5LKyvyWUWwqbjaQfKFXoo8mch.jpg"

        # Set Description & Overview
        if vid_id in descriptions_map:
            video["description"] = descriptions_map[vid_id]
            video["overview"] = descriptions_map[vid_id]
            desc_count += 1

    print("4. Saving fully assembled JSON...")
    with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

    print(f"\nDone! Saved to {OUTPUT_JSON}.")
    print(f"Descriptions matched and injected: {desc_count}")

if __name__ == "__main__":
    main()
