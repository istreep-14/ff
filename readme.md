## NFL Projections Aggregator (Google Sheets + Apps Script)

This Google Apps Script helps you aggregate NFL player projections from multiple sources (CSV/JSON/Sheets), apply your league scoring and roster settings, and calculate fantasy points, VOR (Value Over Replacement) and VOLS (Value Over Last Starter). It writes the consolidated data to a `Players` sheet.

### What you get
- Custom menu in the spreadsheet: Setup Sheets, Refresh Data, Compute Rankings, Run All
- `Sources` sheet to configure multiple input sources (CSV/JSON/Sheet)
- `Settings` sheet with sensible defaults for teams, roster, and scoring
- `Players` output with projections, FantasyPoints, RankOverall, RankPos, VOR, VOLS
- `Logs` sheet for basic status and errors

### Install
1. Create a new Google Spreadsheet.
2. Open Extensions → Apps Script.
3. In the editor, create a file `Code.gs` and paste the contents of `apps-script/Code.gs` from this repo.
4. Save. Return to the spreadsheet and reload to see the “NFL Projections” menu.

Note: If fetching external URLs, you may need to authorize the script (UrlFetchApp).

### Configure Sources
Open the `Sources` sheet and add rows for the sources you have access to.

Headers:
- Enabled: true/false
- SourceName: Friendly name
- Type: CSV, JSON, or SHEET
- URL_or_SpreadsheetID: For CSV/JSON, the URL to the file or API; for SHEET, the spreadsheet ID
- Range_or_JSONPath:
  - CSV: leave blank
  - JSON: optional JSONPath like `$.data` for arrays nested under `data`
  - SHEET: A1 range like `Sheet1!A1:Z`

Important: Ensure your sources provide standard headers, or use the provided header synonyms, or rename the columns in your own source sheet. Standard fields expected:
`Player, Team, Pos, ByeWeek, PassYds, PassTD, Int, RushAtt, RushYds, RushTD, Targets, Rec, RecYds, RecTD, Fumbles, TwoPt, ReturnTD`

### Configure Settings
Open the `Settings` sheet and adjust values as needed:
- League: `NumTeams`, `RosterSize`, `Starters_*`, `Bench_*`, `FLEX`
- Scoring: `PPR`, yardage points per yard, TD points, turnovers, etc.
- VOR baseline: `UseCustomVorRanks` and optional `VOR_Rank_*` overrides
- VOLS baseline: `VOL_UseAuto` or set `VOL_Rank_*` ranks

### Usage
1. From the menu “NFL Projections”, click “Setup Sheets” on first run.
2. Configure `Sources` and `Settings`.
3. Click “Refresh Data” to fetch and aggregate projections into `Players`.
4. Click “Compute Rankings” to calculate FantasyPoints, ranks, VOR, and VOLS.
5. Or use “Run All” to do steps 3 and 4 together.

### Notes
- Only include sources you are permitted to use. Do not scrape or bypass gated content.
- For JSON APIs that require keys, you can store keys in Script Properties and construct URLs accordingly.
- The simple JSONPath supports `$.prop.subprop`. For complex responses, pre-process data into your own Google Sheet and use Type=SHEET.

### Troubleshooting
- Players not appearing: ensure `Enabled=true` for at least one source and that it returns rows with `Player` and `Pos`.
- Incorrect points: verify scoring settings in `Settings`.
- VOR/VOLS seem off: adjust `UseCustomVorRanks`/`VOL_UseAuto` or set explicit ranks per position.
- HTTP errors: the URL may be blocked or require authentication.

### Extending
- Add additional fields to `getStandardFields()` and update scoring.
- Extend `getHeaderSynonyms()` to map alternative column names.
- Add additional positions (e.g., K, DST) and scoring rules if desired.
