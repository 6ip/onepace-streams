"""Print a commit message describing the staged changes.

Run after `git add`, before `git commit`:
    python -u report.py > msg.txt || echo "[auto] data refreshed" > msg.txt
    git commit -F msg.txt

Reporting must never block the commit, so the caller keeps a fallback message.
"""
import json
import os
import subprocess
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MAX_LISTED = 10
MAX_NAMED = 3


def git(*args):
    out = subprocess.run(
        ["git", *args], cwd=BASE_DIR, capture_output=True, text=True, encoding="utf-8"
    )
    return out.stdout if out.returncode == 0 else ""


def load_prefix_map():
    try:
        with open(os.path.join(BASE_DIR, "config.json"), "r", encoding="utf-8") as f:
            return json.load(f).get("PREFIX_MAP", {})
    except (OSError, ValueError):
        return {}


PREFIX_MAP = load_prefix_map()


def human(vid_id):
    """'EH_22' -> 'Egghead 22'. Falls back to the id when the arc is unknown."""
    head, _, tail = vid_id.rpartition("_")
    if head in PREFIX_MAP and tail.isdigit():
        return f"{PREFIX_MAP[head]} {int(tail)}"
    return vid_id


def join_capped(ids):
    shown = [human(i) for i in ids[:MAX_LISTED]]
    extra = len(ids) - len(shown)
    return ", ".join(shown) + (f" +{extra} more" if extra else "")


def hashes(blob):
    try:
        return {s.get("infoHash") for s in json.loads(blob).get("streams") or []}
    except (ValueError, AttributeError, TypeError):
        return set()


def staged_streams():
    """New files, new torrents, and files where only the details changed."""
    added, rereleased, updated = [], [], []
    for line in git("diff", "--cached", "--name-status", "--", "stream/").splitlines():
        parts = line.split("\t")
        if len(parts) < 2 or not parts[-1].endswith(".json"):
            continue
        path = parts[-1]
        vid_id = os.path.basename(path)[:-5]
        if parts[0].startswith("A"):
            added.append(vid_id)
        elif parts[0].startswith("M"):
            # Same torrent means a detail was corrected, not a new release.
            same = hashes(git("show", f"HEAD:{path}")) == hashes(git("show", f":{path}"))
            (updated if same else rereleased).append(vid_id)
    return sorted(added), sorted(rereleased), sorted(updated)


def videos(blob):
    try:
        return {v["id"]: v for v in json.loads(blob)["meta"]["videos"]}
    except (ValueError, KeyError, TypeError):
        return {}


def staged_meta(skip=()):
    """Field change counts, and which fields changed on which episode."""
    added = removed = 0
    fields, touched = {}, {}
    for line in git("diff", "--cached", "--name-only", "--", "meta/").splitlines():
        if not line.endswith(".json"):
            continue
        before = videos(git("show", f"HEAD:{line}"))
        after = videos(git("show", f":{line}"))
        if not after:
            continue
        added += len(set(after) - set(before))
        removed += len(set(before) - set(after))
        for vid_id in set(before) & set(after):
            if vid_id in skip:
                continue
            for key in set(before[vid_id]) | set(after[vid_id]):
                # firstAired and overview mirror released and description.
                if key in ("firstAired", "overview"):
                    continue
                if before[vid_id].get(key) != after[vid_id].get(key):
                    fields[key] = fields.get(key, 0) + 1
                    touched.setdefault(vid_id, set()).add(key)
    return added, removed, fields, touched


LABEL = {
    "released": "release date", "runtime": "runtime", "title": "title",
    "thumbnail": "thumbnail", "description": "description",
    "episode": "episode number", "season": "season",
}


def meta_summary(added, removed, fields, touched):
    # Few episodes: name them. Many: fall back to counts.
    if touched and len(touched) <= MAX_NAMED and not added and not removed:
        return ", ".join(
            f"{human(v)} ({', '.join(LABEL.get(k, k) for k in sorted(keys))})"
            for v, keys in sorted(touched.items())
        )

    bits = []
    if added:
        bits.append(f"+{added} episode" + ("s" if added > 1 else ""))
    if removed:
        bits.append(f"-{removed} episode" + ("s" if removed > 1 else ""))
    for key, count in sorted(fields.items(), key=lambda kv: -kv[1]):
        name = LABEL.get(key, key)
        bits.append(f"{count} {name}" + ("s" if count > 1 else ""))
    return ", ".join(bits)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    new, rereleased, updated = staged_streams()
    # A new episode obviously gains a title and a date, so don't repeat it.
    m_added, m_removed, m_fields, m_touched = staged_meta(skip=set(new))
    meta_bits = meta_summary(m_added, m_removed, m_fields, m_touched)

    head = []
    if new:
        head.append(human(new[0]) + " added" if len(new) == 1 else f"{len(new)} new episodes")
    if rereleased:
        head.append(f"{len(rereleased)} re-released")
    if updated:
        alone = len(updated) == 1 and not new and not rereleased
        head.append(human(updated[0]) + " updated" if alone else f"{len(updated)} updated")
    if not head:
        if len(m_touched) == 1:
            head.append(human(next(iter(m_touched))) + " updated")
        elif m_touched:
            head.append(f"{len(m_touched)} episodes updated")
        else:
            head.append("metadata updated" if meta_bits else "data refreshed")

    lines = [f"[auto] {', '.join(head)}"]
    body = []
    if new:
        body.append(f"New:         {join_capped(new)}")
    if rereleased:
        body.append(f"Re-released: {join_capped(rereleased)}")
    if updated:
        body.append(f"Updated:     {join_capped(updated)}")
    if meta_bits:
        body.append(f"Meta:        {meta_bits}")
    if body:
        lines += ["", *body]

    sys.stdout.write("\n".join(lines) + "\n")


if __name__ == "__main__":
    main()
