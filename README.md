# Google-Sheets-and-Calendar-Sync

## What it does

* Copies an institute timetable (Google Sheet) into your own sheet as **MyEvents**.
* Skips **red cells** (cancellations) and filters by your **allow-list**.
* Preserves headers (D1/E1/…) and adds them into the event title/notes.
* Creates/updates Google Calendar events, **no duplicates**.
* Auto-removes events if they’re marked red or disappear from the sheet.

---

## Setup

1. Open your Google Sheet → **Extensions → Apps Script**.
2. Paste the full script.
3. In the `CFG` block at the top, update:

   * `SOURCE_SPREADSHEET_ID_OR_URL` → institute timetable link/ID
   * `SOURCE_SHEET_NAME` → timetable tab name
   * `EVENT_COLS` → columns with events (e.g. `[3,4,5,6,7,8,9,10]`)
   * `EVENT_HEADER_ROW_INDEX`, `DATETIME_HEADER_ROW_INDEX`, `DATA_START_ROW_INDEX` → row numbers in source for headers/date/time/start of data
   * `CALENDAR_ID` → `"primary"` or your calendar’s ID
   * `TITLE_TEMPLATE` → e.g. `"{COL} - {event}"` for “D1 - ER (FIN)”
4. Add a sheet named **Allowlist** in your file, column A = courses you want kept.
5. Save → in Apps Script editor, run **`setupAll()`** once (authorize when prompted).

---

## Use

* **Manual sync:** Sheet menu → **Sync → Run pipeline now**.
* **Auto sync:** Sheet menu → **Sync → Install time trigger** (runs every X minutes as set in config).
* Events update automatically: red = removed, new events = added, changes = updated.

