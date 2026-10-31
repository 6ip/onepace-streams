<p align="center">
  <img src="https://cdn.jsdelivr.net/gh/6ip/onepace-assets-prm@main/public/logo_o.png" width="110" alt="One Pace Addon">
</p>

# onepace-streams

The data behind the **One Pace Addon** for Stremio: the stream index, the metadata and the catalogs the addon reads at runtime.

**[Install the addon](https://onepace.e6ip.com)**, or find it on [Stremio Addons](https://stremio-addons.net/addons/one-pace-rd).

## What's in here

```
config.json                  arc -> prefix map, shared by both scripts
catalog/                     catalogs served to the addon
meta/                        one metadata file per edit
    pp_onepacee.json           One Pace
    pp_muhnpace.json           Muhn Pace (English dub)
    pp_onipace.json            Onigashima Paced
    pp_KUMA_SHAVED.json        Shaved Egghead & Kuma Cut
stream/                      one file per episode
    *.json                     One Pace
    Muhn/ ONIG/ KUMA_SHAVED/   the alternative edits
    Specials/                  maintained by hand
```

## How it works

- `scr.py` reads the torrent sources from `one_pace.xlsx` and resolves each episode to a magnet info hash. It tries a direct link first, then the release feed, then the arc batch torrents, then a CRC32 search, then a single-episode search. `tracker.json` remembers what resolved, so unchanged episodes are skipped.
- `epis.py` builds `meta/pp_onepacee.json` from the One Pace spreadsheet, with episode titles, descriptions, ratings, specials and thumbnails.
- `report.py` writes the commit message the Action uses.

`config.json` holds the arc names and their prefix codes (for example `alabasta` maps to `AL`), plus `TOTAL_SEASONS`.

## Generated, not hand-edited

A GitHub Action rebuilds `stream/` and `meta/` every three hours. Anything edited there by hand is overwritten on the next run. Send pull requests for the scripts, and open an issue for the data.

## One Pace and the alternative edits

The main catalog comes from [the One Pace project](https://onepace.net/), a volunteer effort that re-cuts One Piece to follow the manga without the filler.

The other edits are made by community members:

- **Muhn Pace** by Muhny
- **Onigashima Paced** by Fedew
- **Shaved Egghead** by Kanes31 and **Kuma Cut** by Rai

They are not part of the main One Pace releases.

[Website](https://onepace.net/) · [r/onepace](https://www.reddit.com/r/onepace/) · [One Pace Discord](https://discord.gg/pacing)

## Credits

- Episode titles from [one-pace-public-subtitles](https://github.com/one-pace/one-pace-public-subtitles)
- Descriptions from the community [episode descriptions sheet](https://docs.google.com/spreadsheets/d/1M0Aa2p5x7NioaH9-u8FyHq6rH3t5s6Sccs8GoC6pHAM/edit)
- Thumbnails from [The Movie Database (TMDB)](https://www.themoviedb.org/). This product uses the TMDB API but is not endorsed or certified by TMDB.
- Episode ratings from the [IMDb Non-Commercial Datasets](https://developer.imdb.com/non-commercial-datasets/), mapped from each One Pace episode to the One Piece episodes it covers. IMDb data is for personal and non-commercial use under IMDb's terms.
- [fedew04/OnePaceStremio](https://github.com/fedew04/OnePaceStremio): the metadata structure this started from

## License

MIT covers `scr.py`, `epis.py`, `report.py` and the rest of the tooling here.

It does **not** cover the data. The episode titles, descriptions and release names are the work of the One Pace project, the editors of the alternative edits, and their contributors. This repository only indexes and reformats them so the addon can read them.