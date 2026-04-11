import urllib.request
import json
import re
import urllib.parse
import os
import time
import pysubs2

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

# --- Regex Patterns ---
_KARAOKE_PATTERN = re.compile(r'(karaoke|kara|romaji|rom|kanji|furigana|credits?)')
_OP_ED_STYLE_PATTERN = re.compile(r'(op\d*|ed\d*|ending|opening|song)')
_X_POS_PATTERN = re.compile(r'\\(?:pos|move)\(([-+]?\d*\.?\d+)')
_Y_POS_PATTERN = re.compile(r'\\(?:pos|move)\s*\([-+]?\d*\.?\d+\s*,\s*([-+]?\d*\.?\d+)')

def ms_to_vtt_time(ms: int) -> str:
    """Helper to convert integer milliseconds back to WEBVTT timestamp formatting."""
    h = ms // 3600000
    ms %= 3600000
    m = ms // 60000
    ms %= 60000
    s = ms // 1000
    ms %= 1000
    return f"{h:02d}:{m:02d}:{s:02d}.{ms:03d}"

def fix_rtl_visual_typing(text: str) -> str:
    """
    Fixes Arabic punctuation typed visually and enforces RTL per line.
    Correctly maps Aegisub's visual left (index 0) to Arabic's logical end.
    """
    flip_map = str.maketrans('«»()[]{}', '»«)(][}{')
    enclosing_chars = r'"\'«»()[]{}'
    terminal_chars = r'!؟?.,،؛:'
    
    fixed_lines = []
    for line in text.split('\n'):
        # 1. Clean existing BiDi markers
        line = re.sub(r'[\u200e\u200f\u202a\u202b\u202c\u202d\u202e]', '', line)
        
        if not line.strip():
            fixed_lines.append(line)
            continue
            
        core_text = line.strip()
        
        # 2. Extract ALL punctuation from the LEFT side.
        # In Aegisub, the left side is the visual left. For Arabic, this is the logical END.
        left_enc = ""
        left_term = ""
        while core_text and core_text[0] in enclosing_chars + terminal_chars:
            if core_text[0] in terminal_chars:
                left_term += core_text[0]
            else:
                left_enc += core_text[0]
            core_text = core_text[1:].lstrip()
            
        # 3. Extract ONLY enclosures from the RIGHT side.
        # In Aegisub, the right side is the visual right (logical START).
        # Terminals are purposely ignored here to fix inline typos correctly.
        end_enc = ""
        while core_text and core_text[-1] in enclosing_chars:
            end_enc = core_text[-1] + end_enc
            core_text = core_text[:-1].rstrip()
            
        # 4. Smart Swap Logic
        needs_swap = False
        if left_enc and end_enc:
            # Both sides have quotes -> Typical LTR backward typing.
            needs_swap = True
        elif end_enc and not left_enc:
            # Only end quote. Is it an orphan?
            for char in end_enc:
                if char in core_text:
                    needs_swap = True
                    break
        elif left_enc and not end_enc:
            needs_swap = True
            
        if needs_swap:
            new_start_enc = end_enc.translate(flip_map)
            new_end_enc = left_enc.translate(flip_map)
        else:
            new_start_enc = left_enc
            new_end_enc = end_enc
            
        # 5. Fix inline visual typos
        # NEW FIX: Require at least one alphanumeric character ([^\W_]) inside the enclosures!
        # This prevents the regex from aggressively matching the space/punctuation between two separate valid quotes.
        core_text = re.sub(r'»([^«»]*[^\W_][^«»]*)«', r'«\1»', core_text) # Fixes »word«
        core_text = re.sub(r'\)([^()]*[^\W_][^()]*)\(', r'(\1)', core_text) # Fixes )word(
        core_text = re.sub(r'\]([^\[\]]*[^\W_][^\[\]]*)\[', r'[\1]', core_text) # Fixes ]word[
        core_text = re.sub(r'\}([^{}]*[^\W_][^{}]*)\{', r'{\1}', core_text) # Fixes }word{
        
        # 6. Reconstruct line securely
        fixed_line = f"{new_start_enc}{core_text}{left_term}{new_end_enc}"
        fixed_lines.append(f"\u202B{fixed_line}\u202C")
        
    return '\n'.join(fixed_lines)
    
def process_op_ed_file(ass_content: str, offset_ms: int, lang_code: str) -> list:
    # Load the ASS file directly from the downloaded string
    subs = pysubs2.SSAFile.from_string(ass_content)
    raw_dialogues = []
    sync_ms = None

    # 1. First pass: find the sync time
    for line in subs:
        if line.name.lower() == "sync":
            sync_ms = line.start
            break

    # 2. Second pass: process lines
    for line in subs:
        name = line.name.lower()
        effect = line.effect.lower()
        style = line.style.lower()
        text_raw = line.text

        # Skip sync lines in this loop, or scene ends
        if name == "sync":
            continue
        
        clean_text_pre = line.plaintext.lower()
        if "scene ends" in clean_text_pre or name in ["op", "ed", "ending"]:
            continue

        # --- 1. FILTERS ---
        if _KARAOKE_PATTERN.search(style):
            continue
        if r"\p1" in text_raw or r"\p2" in text_raw or r"\p4" in text_raw or r"\p0" in text_raw:
            continue

        if not line.is_comment:
            if r"\k" in text_raw.lower() or name in ["lead-in", "hi-light", "verse", "karaoke", "mask", "glow", "shape", "gradient", "dust", "petals", "border clip", "move", "circle", "cross"]:
                continue
            if lang_code != "ara" and ("fx" in effect or "effector" in effect or "kara effector" in text_raw.lower()):
                continue

        # --- EXTRACT X POS (From raw text containing tags) ---
        x_pos = 0.0
        pos_match = _X_POS_PATTERN.search(text_raw)
        if pos_match:
            try:
                x_pos = float(pos_match.group(1))
            except ValueError:
                pass

        # --- 2. CLEAN TEXT ---
        # pysubs2 strips {tags} but leaves \N and \h. We convert them to standard spaces/newlines.
        clean_text = line.plaintext.replace(r"\h", " ").replace("\\h", " ")
        clean_text = clean_text.replace("\\N", "\n").replace("\\n", "\n")
        clean_text = re.sub(r'[\u200e\u200f\u202a\u202b\u202c\u202d\u202e]', '', clean_text)

        # FIX: Check length WITHOUT tatweel and harakat to catch fragments like "ـيـ"
        test_text = re.sub(r'[\u0640\u064B-\u065F\u0670]', '', clean_text)
        if not re.search(r'[^\W_]', test_text):
            continue

        clean_no_marks = test_text.strip().lower()
        if len(clean_no_marks) == 1:
            valid_singles = "aioyeuàôو"
            if clean_no_marks not in valid_singles and not clean_no_marks.isdigit():
                continue
            if clean_no_marks.isdigit() and ("fx" in effect or "kara" in style or "title" in style or "sign" in style or name in ["op", "ed"]):
                continue

        if "code" in effect or "template" in effect or "fxgroup" in clean_text.lower() or "_g." in clean_text.lower() or "retime" in clean_text.lower():
            continue

        clean_lower = clean_text.strip().lower()
        if clean_lower == style or clean_lower == "roger monologue":
            continue
        if re.fullmatch(r"(op|ed)[ -][a-z]+", clean_lower) or re.fullmatch(r"[a-z]+[ -](op|ed)", clean_lower):
            continue
        if re.search(r'^[-=]{3,}.*[-=]{3,}$', clean_lower) or re.fullmatch(r'[-=\s]+', clean_lower):
            continue
        if clean_lower.strip(' -=_') in ["ending", "opening", "op", "ed", "dialogue", "credits", "title", "signs"]:
            continue

        if clean_text == "":
            clean_text = " "

        if lang_code == "ara" and clean_text != " " and not re.search(r'[\u0600-\u06FF]', clean_text):
            continue

        if re.search(r'[\u0600-\u06FF]', clean_text):
            clean_text = clean_text.strip()
            if lang_code == "ara":
                clean_text = fix_rtl_visual_typing(clean_text)

        raw_dialogues.append({
            "raw_start": line.start,
            "raw_end": line.end,
            "text": clean_text,
            "style": style,
            "x_pos": x_pos
        })
        
    if not raw_dialogues:
        return []
        
    deduped = []
    seen = set()
    for d in raw_dialogues:
        identifier = (d["raw_start"], d["raw_end"], d["text"], d["style"])
        if identifier not in seen:
            seen.add(identifier)
            deduped.append(d)
    raw_dialogues = deduped

    if sync_ms is not None:
        base_ms = sync_ms
    else:
        base_ms = min(d["raw_start"] for d in raw_dialogues)

    dialogues = []
    for d in raw_dialogues:
        dialogues.append({
            "start_ms": (d["raw_start"] - base_ms) + offset_ms,
            "end_ms": (d["raw_end"] - base_ms) + offset_ms,
            "text": d["text"],
            "style": d["style"],
            "x_pos": d["x_pos"]
        })

    dialogues.sort(key=lambda x: (x["start_ms"], x["end_ms"]))
    
    # --- NEW: Active Clusters Algorithm (Handles Simultaneous Tracks) ---
    active_clusters = []
    
    for d in dialogues:
        placed = False
        for cluster in active_clusters:
            prev = cluster[0]
            if d["style"] == prev["style"] and abs(d["start_ms"] - prev["start_ms"]) < 1500 and abs(d["end_ms"] - prev["end_ms"]) < 1500:
                cluster.append(d)
                placed = True
                break
        
        if not placed:
            active_clusters.append([d])
            
    clustered_dialogues = []
    
    for cluster in active_clusters:
        # Spatial Sorting: Descending for RTL (Arabic), Ascending for LTR
        if lang_code == "ara":
            cluster.sort(key=lambda x: x["x_pos"], reverse=True)
        else:
            cluster.sort(key=lambda x: x["x_pos"], reverse=False)

        unique_parts = []
        seen_parts = []
        for x in cluster:
            is_layer_dup = False
            for seen_text, seen_x in seen_parts:
                if x["text"] == seen_text and abs(x["x_pos"] - seen_x) < 5.0:
                    is_layer_dup = True
                    break
            if not is_layer_dup:
                seen_parts.append((x["text"], x["x_pos"]))
                unique_parts.append(x)

        parts = [x["text"] for x in unique_parts]
        valid_parts = [p for p in parts if p.strip()]
        if valid_parts:
            avg_len = sum(len(p) for p in valid_parts) / len(valid_parts)
            separator = "" if avg_len <= 1.5 else " "
            merged_text = separator.join(parts)
            merged_text = re.sub(r'\s+', ' ', merged_text).strip()
            
            clustered_dialogues.append({
                "start_ms": min(x["start_ms"] for x in cluster),
                "end_ms": max(x["end_ms"] for x in cluster),
                "text": merged_text
            })

    # --- 5. DEDUPLICATE OVERLAPPING FX LAYERS ---
    final_dialogues = []
    for d in clustered_dialogues:
        if not d["text"]: continue
        is_dup = False
        clean_d = re.sub(r'[^\w]', '', d["text"])
        for f in final_dialogues:
            clean_f = re.sub(r'[^\w]', '', f["text"])
            if clean_d and clean_d == clean_f and abs(f["start_ms"] - d["start_ms"]) < 4000:
                is_dup = True
                f["end_ms"] = max(f["end_ms"], d["end_ms"])
                f["start_ms"] = min(f["start_ms"], d["start_ms"])
                break
        if not is_dup:
            if re.search(r'[\u0600-\u06FF]', d["text"]):
                clean_d_text = d["text"].replace("\u202B", "").replace("\u202C", "").strip()
                if lang_code == "ara":
                    d["text"] = fix_rtl_visual_typing(clean_d_text)
                else:
                    d["text"] = f"\u202B{clean_d_text}\u202C"
            final_dialogues.append(d)

    return final_dialogues

def ass_to_vtt(ass_content: str, op_dialogues: list = None, ed_dialogues: list = None, lang_code: str = "eng") -> str:
    subs = pysubs2.SSAFile.from_string(ass_content)
    
    op_start_ms = None  
    ed_start_ms = None  
    dialogues = []

    for line in subs:
        name = line.name.lower()
        style = line.style.lower()
        effect = line.effect.lower()
        text_raw = line.text

        # Handle Sync points for OP/ED
        if line.is_comment:
            if name == "op":
                op_start_ms = line.start
            elif name in ["ed", "ending"] or "ending" in style:
                ed_start_ms = line.start
            
            if not _OP_ED_STYLE_PATTERN.search(style):
                continue

        if _KARAOKE_PATTERN.search(style):
            continue
            
        if r"\k" in text_raw.lower() or name in ["lead-in", "hi-light", "verse", "karaoke", "mask", "glow", "shape", "gradient", "dust", "petals", "border clip", "move", "circle", "cross"]:
            continue

        if lang_code != "ara" and ("fx" in effect or "effector" in effect or "kara effector" in text_raw.lower()):
            continue
            
        if lang_code == "spa" and re.search(r'(main|flashback|thought|secondary|caption|title)', style):
            continue

        # Extract X and Y pos from RAW text before tags are stripped
        x_pos = 0.0
        pos_match_x = _X_POS_PATTERN.search(text_raw)
        if pos_match_x:
            try: x_pos = float(pos_match_x.group(1))
            except ValueError: pass

        y_pos = 1000.0
        pos_match_y = _Y_POS_PATTERN.search(text_raw)
        if pos_match_y:
            try: y_pos = float(pos_match_y.group(1))
            except ValueError: pass
                
        dialogues.append((line, x_pos, y_pos))

    processed_dialogues = []
    
    # PASS 1: Extract and clean raw components
    for line, x_pos, y_pos in dialogues:
        text_raw = line.text
        
        # Block drawings
        if r"\p1" in text_raw or r"\p2" in text_raw or r"\p4" in text_raw:
            continue

        # Get clean text natively
        text = line.plaintext.replace(r"\h", " ").replace("\\h", " ")
        text = text.replace("\\N", "\n").replace("\\n", "\n")
        text = re.sub(r'[\u200e\u200f\u202a\u202b\u202c\u202d\u202e]', '', text).strip('\r\n\t')

        if text == "" or "mpv.io" in text.lower() or "mpvio" in text.lower():
            continue
          
        test_text = re.sub(r'[\u0640\u064B-\u065F\u0670]', '', text)
        if not re.search(r'[^\W_]', test_text):
            continue
            
        if "code" in line.effect.lower() or "template" in line.effect.lower() or "fxgroup" in text_raw.lower() or "_g." in text_raw.lower() or "retime" in text_raw.lower():
            continue
            
        clean_lower = text.strip().lower()
        if clean_lower == line.style.lower() or clean_lower == "roger monologue":
            continue
        if re.fullmatch(r"(op|ed)[ -][a-z]+", clean_lower) or re.fullmatch(r"[a-z]+[ -](op|ed)", clean_lower):
            continue
        if re.search(r'^[-=]{3,}.*[-=]{3,}$', clean_lower) or re.fullmatch(r'[-=\s]+', clean_lower):
            continue
        if clean_lower.strip(' -=_') in ["ending", "opening", "op", "ed", "dialogue", "credits", "title", "signs"]:
            continue

        processed_dialogues.append({
            "start_ms": line.start,
            "end_ms": line.end,
            "text": text.strip(),
            "style": line.style.lower(),
            "x_pos": x_pos,
            "y_pos": y_pos,
            "effect": line.effect.lower()
        })

    # PASS 2: Spatial Clustering (Stitches fragments back into words)
    active_clusters = []
    for d in processed_dialogues:
        placed = False
        for cluster in active_clusters:
            prev = cluster[0]
            # FIX: Only cluster if they START at exactly the same time (within 200ms).
            # This prevents sequential dialogue lines from merging!
            time_match = abs(d["start_ms"] - prev["start_ms"]) < 200 and abs(d["end_ms"] - prev["end_ms"]) < 200
            
            if d["style"] == prev["style"] and time_match and abs(d["y_pos"] - prev["y_pos"]) < 100:
                cluster.append(d)
                placed = True
                break
        if not placed:
            active_clusters.append([d])

    clustered_dialogues = []
    for cluster in active_clusters:
        # Sort spatially based on language
        if lang_code == "ara":
            cluster.sort(key=lambda x: x["x_pos"], reverse=True)
        else:
            cluster.sort(key=lambda x: x["x_pos"], reverse=False)

        unique_parts = []
        seen_parts = []
        for x in cluster:
            is_layer_dup = False
            for seen_text, seen_x in seen_parts:
                if x["text"] == seen_text and abs(x["x_pos"] - seen_x) < 5.0:
                    is_layer_dup = True
                    break
            if not is_layer_dup:
                seen_parts.append((x["text"], x["x_pos"]))
                unique_parts.append(x)

        parts = [x["text"] for x in unique_parts]
        valid_parts = [p for p in parts if p.strip()]
        if not valid_parts:
            continue

        # If pieces are small, join with empty string to rebuild word. Otherwise space.
        avg_len = sum(len(re.sub(r'[\u0640\u064B-\u065F\u0670]', '', p)) for p in valid_parts) / len(valid_parts)
        separator = "" if avg_len <= 3.0 else " "
        
        merged_text = separator.join(parts)
        merged_text = re.sub(r'\s+', ' ', merged_text).strip()
        
        # Final noise filter (applied AFTER words are stitched back together)
        clean_no_marks = re.sub(r'[\u0640\u064B-\u065F\u0670]', '', merged_text).strip().lower()
        if len(clean_no_marks) == 1:
            valid_singles = "aioyeuàôو"
            style = cluster[0]["style"]
            effect = cluster[0]["effect"]
            if clean_no_marks not in valid_singles and not clean_no_marks.isdigit():
                continue
            if clean_no_marks.isdigit() and ("fx" in effect or "kara" in style or "title" in style or "sign" in style):
                continue
                
        # Strip cosmetic Arabic tatweels and fix visual typing
        if lang_code == "ara":
            merged_text = merged_text.replace("ـ", "")
            if re.search(r'[\u0600-\u06FF]', merged_text):
                merged_text = fix_rtl_visual_typing(merged_text)

        style = cluster[0]["style"]
        if re.search(r'(title|caption|sign)', style):
            merged_text = f"<b>{merged_text}</b>"

        clustered_dialogues.append({
            "start_ms": min(x["start_ms"] for x in cluster),
            "end_ms": max(x["end_ms"] for x in cluster),
            "text": merged_text,
            "y_pos": sum(x["y_pos"] for x in cluster) / len(cluster) # Average vertical height
        })

    # Add Injected OP/ED Themes
    if op_dialogues and op_start_ms is not None:
        for op_line in op_dialogues:
            clustered_dialogues.append({
                "start_ms": op_line["start_ms"] + op_start_ms,
                "end_ms": op_line["end_ms"] + op_start_ms,
                "text": op_line["text"],
                "y_pos": 1000.0
            })
            
    if ed_dialogues and ed_start_ms is not None:
        for ed_line in ed_dialogues:
            clustered_dialogues.append({
                "start_ms": ed_line["start_ms"] + ed_start_ms,
                "end_ms": ed_line["end_ms"] + ed_start_ms,
                "text": ed_line["text"],
                "y_pos": 1000.0
            })

    # PASS 3: Temporal Merge (Flattens rapidly flickering identical text frames)
    clustered_dialogues.sort(key=lambda x: (x["start_ms"], x["text"]))
    
    merged_dialogues = []
    for d in clustered_dialogues:
        placed = False
        for m in reversed(merged_dialogues):
            if d["start_ms"] - m["end_ms"] > 1500:
                break
                
            if d["text"] == m["text"] and d["start_ms"] <= m["end_ms"] + 500:
                m["end_ms"] = max(m["end_ms"], d["end_ms"])
                m["start_ms"] = min(m["start_ms"], d["start_ms"])
                placed = True
                break
                
        if not placed:
            merged_dialogues.append(d)

    processed_dialogues = merged_dialogues

    # PASS 4: VTT Output Generation
    vtt_lines = [
        "WEBVTT",
        "",
        "STYLE",
        "::cue(c.color9CD5FF) { color: #9CD5FF; }",
        "::cue(c.colora8c7fa) { color: #a8c7fa; }",
        "",
        "1",
        "00:00:01.000 --> 00:00:07.000 line:5% align:middle",
        "<b><c.color9CD5FF>One Pace Premium</c></b>",
        "Keep the project alive: <c.colora8c7fa>ko-fi.com/not6ip</c>",
        ""
    ]
    
    grouped_dialogues = {}
    for d in processed_dialogues:
        time_key = (d["start_ms"], d["end_ms"])
        if time_key not in grouped_dialogues:
            grouped_dialogues[time_key] = []
            
        if not any(item["text"] == d["text"] for item in grouped_dialogues[time_key]):
            grouped_dialogues[time_key].append(d)

    counter = 2
    for (start, end), items in grouped_dialogues.items():
        items.sort(key=lambda x: x["y_pos"])
        texts = [x["text"] for x in items]
        
        vtt_lines.append(str(counter))
        vtt_lines.append(f"{ms_to_vtt_time(start)} --> {ms_to_vtt_time(end)}")
        vtt_lines.append("\n".join(texts))
        vtt_lines.append("")
        counter += 1

    return "\n".join(vtt_lines)

def clean_string(s):
    return re.sub(r'[\d\s\-]', '', s).lower()

def parse_properties_rules(prop_text: str) -> list:
    """Parses the sub.properties file into a queryable list of rules."""
    rules = []
    for line in prop_text.split('\n'):
        line = line.strip()
        if not line or line.startswith('#'): continue
        if '.OP' not in line and '.ED' not in line: continue
        
        left, _, path = line.partition('=')
        if not path: continue
        
        # Fix the dbf typo in their properties file
        left = left.replace('{01.06}', '{01..06}')
        
        key_part, op_ed_part = left.rsplit('.', 1)
        is_op = op_ed_part.startswith('OP')
        op_type = "OP" if is_op else "ED"
        
        lang_suffix = op_ed_part[2:].strip('_') # e.g. "de", "ar"
        if lang_suffix in LANG_MAP:
            lang_suffix = LANG_MAP[lang_suffix]
        elif not lang_suffix: 
            lang_suffix = "eng"
            
        rules.append({
            "pattern": key_part, 
            "type": op_type, 
            "lang": lang_suffix, 
            "path": path.strip()
        })
    return rules

def match_rule(arc: str, ep: int, pattern: str) -> bool:
    """Uses Regex to match an arc and episode number against bash-style brace expansion."""
    def expand_range(m):
        start, end = int(m.group(1)), int(m.group(2))
        width = len(m.group(1)) # Keep zero padding (e.g. 01)
        return "(" + "|".join(f"{i:0{width}d}" for i in range(start, end + 1)) + ")"
        
    regex = re.sub(r'\{(\d+)\.\.(\d+)\}', expand_range, pattern)
    regex = re.sub(r'\{([^}]+)\}', lambda m: "(" + m.group(1).replace(',', '|') + ")", regex)
    regex = regex.replace('*', '.*')
    
    target = f"{arc}_{ep:02d}"
    return re.fullmatch(regex, target) is not None

def get_op_ed_paths(arc_key: str, ep_num: int, lang_code: str, rules: list):
    op_path, ed_path = None, None
    for op_type in ["OP", "ED"]:
        best_path = None
        for rule in rules:
            if rule["type"] == op_type and rule["lang"] == lang_code:
                if match_rule(arc_key, ep_num, rule["pattern"]):
                    best_path = rule["path"]
                    break
        if op_type == "OP": op_path = best_path
        else: ed_path = best_path
    return op_path, ed_path

def fetch_op_ed(path: str):
    """Downloads and caches the OP/ED template into themed sub-folders."""
    if not path: return None
    
    parts = path.split('/')
    filename = parts[-1]
    theme_name = parts[-2] if len(parts) >= 2 else "Unknown"
    
    local_path = os.path.join(OUTPUT_SUBS_DIR, "op_ed", theme_name, filename)
    
    if os.path.exists(local_path):
        with open(local_path, 'r', encoding='utf-8') as f:
            return f.read()
            
    print(f"    [+] Downloading Theme: {filename} into {theme_name}/")
    download_url = RAW_ASS_BASE_URL + "main/" + urllib.parse.quote(path)
    try:
        req = urllib.request.Request(download_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as resp:
            content = resp.read().decode('utf-8-sig', errors='ignore')
            
        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        with open(local_path, 'w', encoding='utf-8') as f:
            f.write(content)
        time.sleep(0.5)
        return content
    except Exception as e:
        print(f"    [-] Failed to fetch theme {path}: {e}")
        return None
        
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
    print("[?] Fetching sub.properties for OP/ED mapping...")
    try:
        prop_req = urllib.request.Request(RAW_ASS_BASE_URL + "main/sub.properties", headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(prop_req) as response:
            op_ed_rules = parse_properties_rules(response.read().decode('utf-8'))
    except Exception as e:
        print(f"[-] Failed to fetch sub.properties: {e}")
        op_ed_rules = []

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
        # Explicitly ignore the Release folder
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

        # Fallback for Cover Stories & Specials
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

        if "Cover Stories" in arc_folder or prefix in ["BUGGYS_CREW", "COVER_KOBYMEPPO", "COVER_SHSS"]:
            stremio_id = f"{prefix}_1"
        else:
            stremio_id = f"{prefix}_{ep_num}"

        name_without_ext = filename.rsplit('.', 1)[0]
        ep_str = str(ep_num).zfill(2)
        idx = name_without_ext.rfind(ep_str)
        
        # Smarter Language Extraction
        if idx != -1:
            raw_lang_str = name_without_ext[idx + len(ep_str):].strip()
        else:
            words = name_without_ext.split()
            lang_parts = []
            for w in reversed(words):
                w_lower = w.lower()
                if w_lower in LANG_MAP or w_lower in ['alternate', 'dub', 'cc', 'typesetting']:
                    lang_parts.insert(0, w_lower)
                else:
                    break 
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

        vtt_filename = f"{unique_sub_id}.vtt"
        
        nested_dir = os.path.join(OUTPUT_SUBS_DIR, arc_folder, ep_folder)
        os.makedirs(nested_dir, exist_ok=True) 
        
        local_vtt_path = os.path.join(nested_dir, vtt_filename)
        
        rel_path = f"{arc_folder}/{ep_folder}/{vtt_filename}"
        cdn_url = CDN_SRT_BASE_URL + urllib.parse.quote(rel_path, safe='/')
        
        op_path, ed_path = get_op_ed_paths(arc_key, ep_num, lang_code, op_ed_rules)     
        file_exists = os.path.exists(local_vtt_path)
        
        if file_exists and path not in local_hashes:
            local_hashes[path] = file_sha
            needs_download = False
        else:
            needs_download = not file_exists or (path in local_hashes and local_hashes[path] != file_sha)

        if needs_download:
            download_url = RAW_ASS_BASE_URL + urllib.parse.quote(path)
            
            success = False
            for attempt in range(5):
                try:
                    dl_req = urllib.request.Request(download_url, headers={'User-Agent': 'Mozilla/5.0'})
                    with urllib.request.urlopen(dl_req) as dl_resp:
                        ass_text = dl_resp.read().decode('utf-8-sig', errors='ignore')
                        
                    # Fetch, Parse, and Inject OP/ED
                    op_dialogues, ed_dialogues = None, None
                    if op_path:
                        op_content = fetch_op_ed(op_path)
                        if op_content: op_dialogues = process_op_ed_file(op_content, 0, lang_code)
                    if ed_path:
                        ed_content = fetch_op_ed(ed_path)
                        if ed_content: ed_dialogues = process_op_ed_file(ed_content, 0, lang_code)
                        
                    # Call the VTT function with injected themes
                    vtt_text = ass_to_vtt(ass_text, op_dialogues, ed_dialogues, lang_code)
                    
                    with open(local_vtt_path, "w", encoding="utf-8") as f:
                        f.write(vtt_text)
                        
                    local_hashes[path] = file_sha
                        
                    success = True
                    time.sleep(0.5) 
                    break
                    
                except urllib.error.HTTPError as e:
                    if e.code in [429, 403]:
                        wait_time = 3 ** attempt
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

        subtitles_dict[stremio_id].append({
            "id": unique_sub_id,
            "url": cdn_url,
            "lang": lang_code 
        })

        if i > 0 and i % 50 == 0:
            print(f"    -> Processing {i}/{len(ass_files)} files... (Autosaving hashes)")
            with open(HASHES_FILE, "w", encoding="utf-8") as f:
                json.dump(local_hashes, f, indent=4)

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