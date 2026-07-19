# MGA Results - Project Instructions

## On Every Conversation Start

1. Read `mga_standings.py` and check the `TOURNAMENTS` config list
2. Open `Results.xlsx` and compare sheet names against configured `sheet_name` values
3. If there are NEW sheets not yet in the config, ask the user to confirm:
   - Which tournament in the config does this sheet belong to (match by name/date)
   - Date played (YYYY-MM-DD) - verify against existing config for weather fetch
   - Event type and team size (if not already in config)
   - Number of places paid
   - Flighted? (yes/no)
4. Set the `sheet_name` in the matching `TOURNAMENTS` entry
5. Run `python mga_standings.py` to regenerate outputs
6. If no new sheets found, tell the user "Config is up to date" and ask what they'd like to work on

## Key Files

- `mga_standings.py` - single-file generator: parsing, scoring, analysis, HTML/PDF output
- `Results.xlsx` - tournament results (one sheet per event)
- `mga_standings.html` / `index.html` / `mga_standings.pdf` - generated outputs (index.html is the GitHub Pages copy)

## Folder Layout

- Root - the live build only: `mga_standings.py`, `Results.xlsx`, generated HTML/PDF, `CLAUDE.md`
- `sources/` - raw source inputs kept for manual data entry (bracket JPGs/PDFs, leaderboards, schedule); not read by the script
- `points_structure/` - points-restructure workstream: `build_proposed.py`, `gen_analysis.py`, points reference/proposed xlsx, and the generated `points_analysis` outputs
  - `RYDER CUP POINTS 2026.xlsx` - points structure reference
  - `RYDER CUP POINTS 2026 - PROPOSED.xlsx` - proposed restructure (pending committee approval)
- `archive/` - old-season material (2021/2022, Eldorado FedX) and unused avatar assets

`sources/`, `points_structure/`, and `archive/` are git-ignored (local-only); only the live build is tracked and deployed.

## Tournament Config Format

Each tournament in `TOURNAMENTS` is a 9-element tuple:
```python
(display_name, sheet_name, event_type, team_size, places_paid, has_flights, date_str, participation_pts, season)
```

- `sheet_name = None` means the tournament hasn't been played yet
- `date_str` is `"YYYY-MM-DD"` or `None` - used for weather auto-fetch (Open-Meteo, Eldorado CC, McKinney TX)
- `participation_pts` is always per-player (e.g., 25 for Member-Member, 10 for most others)
- `season` is `"2025-26"` or `"2026-27"`

## Sheet Format Types

### Standard (GolfGenius export)
- `Pos. | Players | To Par | Thru | Total` (or with a GolfGenius link in col A)
- Player names in `"Last, First + Last, First"` format
- Flighted sheets have `"Flight N"` header rows separating sections

### Member-Member (match play)
Auto-detected when column G row 1 = `"Results"`:
- Column A: match play brackets with all participants
- Columns G & H: finishing place + player name in `"First Last"` format
- Names normalized to `"Last, First"` with last-name fallback matching

## Points System

- Placement points from `POINTS_TABLE`, keyed by `(event_type, places_paid)`
- Points source: "points_structure/RYAN PARKS RYDER CUP RECOMMENDATIONS.xlsx" (transcribed into `POINTS_TABLE` in the script; the xlsx is reference only, not read at runtime)
- Participation points: per-player flat amount, awarded to every participant regardless of finish
- Ties within paid places: average points across all positions occupied (e.g. 4-way T1 with 4 places = (100+80+60+40)/4 = 70)
- Ties straddling the last paid place: extend phantom places using the established increment between consecutive places (floored at 0), then average (e.g. 2-way T4 with 4 places = (40+20)/2 = 30)
- Ties entirely outside paid places: 0 points
- All values displayed as per-player only - no team totals or `/pl` caveats

## Places Paid Rule (Standard Events)

For standard events (Individual, 2-Man, 3-Man, 4-Man): target **1/3 of the field** getting performance points.
- Use `calc_places_paid(event_type, total_teams, num_flights)` to pick the closest available option
- Equal payout across all flights - every flight pays the same number of places
- **Minimum 3 places paid**, always
- Available options: Individual/2-Man/3-Man: 3, 4, or 5 | 4-Man: 4, 5, or 6
- Special events (Member/Member, Presidents Cup, Lonely Guy, 2-Man Match Play, Eldo Cup) have places_paid set explicitly - do NOT apply the 1/3 rule

## Schedule Sections

Three sections in Season Schedule & Points:
- **Completed**: tournaments with results, plus cancelled events as `Cancelled*`
- **Ongoing**: active match play brackets with deadline disclaimer
- **Upcoming**: future tournaments

## Standings

- All players in one ranked table - no compact/abbreviated section
- Columns: Rank | Player | [Tournament columns] | Events | Participation Pts | Total Pts
- Footer (tfoot): right-aligned participation pts reference, repeats on every printed page

## Formatting Standards

- No em dashes - use hyphens (-) everywhere
- Page headers: white background, dark green text (#1a472a), bold, green bottom border
- Table headers: light grey background (#f5f5f5), dark green bold text, green bottom border
- Subtitles: full opacity, not greyed out
- Schedule page: must fit on one PDF page
- PDF: Letter size, landscape, 10mm margins, Playwright

## Adding a New Tournament

1. Check the new sheet name in Results.xlsx
2. Match it to the correct unplayed tournament in `TOURNAMENTS` (by name/date)
3. Confirm the date with the user (weather will be fetched for that date)
4. Update `sheet_name` in the config tuple
5. Sheet format is auto-detected by the parser
6. Run `python mga_standings.py` to regenerate
