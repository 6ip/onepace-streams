# onepace-streams

Data backend for the [One Pace Premium](https://onepace-premium.1102011.xyz/) Stremio addon. Holds the stream index, metadata, and catalog files the addon reads at runtime.

## Repository layout

```
config.json          # Arc → prefix map shared across all scripts
├── catalog/         # Catalog JSON files served to the Stremio addon
├── meta/            # Metadata JSONs for each supported edit
│   ├── pp_onepacee.json      # Main One Pace (built by epis.py)
│   ├── pp_muhnpace.json      # Muhn Pace English Dub
│   ├── pp_onipace.json       # Onigashima Paced
│   └── pp_KUMA_SHAVED.json   # Shaved Egghead & Kuma Cut
└── stream/          # Stream JSON files (one per episode)
    ├── *.json             # Standard One Pace — built by scr.py
    ├── ONIG/              # Onigashima Paced
    ├── Muhn/              # Muhn Pace English Dub
    └── KUMA_SHAVED/       # Shaved Egghead & Kuma Cut
```

## Scripts

### `scr.py` — Stream builder

Reads torrent sources from `one_pace.xlsx` and resolves each entry to a magnet info hash, writing output to `stream/`. Uses a fallback chain per episode: direct link → One Pace official site → arc batch torrent → CRC32 search → single episode search.

`tracker.json` tracks the last-known input state per file so unchanged episodes are skipped. A `stream/st_purge.txt` file is written each run listing only changed files for cache invalidation.

### `epis.py` — Metadata builder

Assembles `meta/pp_onepacee.json` from:

- Base structure from [fedew04/OnePaceStremio](https://github.com/fedew04/OnePaceStremio)
- Official episode titles from [one-pace-public-subtitles](https://github.com/one-pace/one-pace-public-subtitles/blob/main/main/title.properties)
- Descriptions from a Google Sheets spreadsheet
- Specials defined in `meta/specials.json`

## Config

`config.json` is the source of truth for arc names and their Stremio prefix codes (e.g. `alabasta → AL`). Both scripts read it at startup. It also defines `TOTAL_SEASONS` used for season poster generation.

## Credits

- [The One Pace Project](https://onepace.net/) — the edits this addon serves
- [fedew04/OnePaceStremio](https://github.com/fedew04/OnePaceStremio) — original Stremio metadata structure
- [one-pace/one-pace-public-subtitles](https://github.com/one-pace/one-pace-public-subtitles) — official title data

If this saved you from managing torrents manually, consider [supporting server costs on Ko-fi](https://ko-fi.com/not6ip).
