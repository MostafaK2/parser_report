# Shift Report → DAR

Turns a filled-in Shift Report (.xlsx) into the written DAR, and lets you say who did each
tour, lock and escort before you copy it.

The Excel file is **only read, never changed**.

## Setup

1. Python 3.8 or newer (python.org installers include tkinter).
2. `pip install -r requirements.txt` (this installs openpyxl).
   - On Linux, tkinter may also need `sudo apt install python3-tk`.

## Using the app

```
python dar_app.py                 # then Open report… (Ctrl+O)
python dar_app.py report.xlsx     # open a report straight away
```

1. **Open report.** The top panels show who is on duty at each site and your settings.
2. **Go through the list.** Each tour, lock and escort is one row:

   | Status | Meaning |
   |---|---|
   | only one on site | One officer/Lt. at that site, so it's filled in for you |
   | 34R Prentiss rule | Filled in by a location rule (Doble Lt., else Disp., else PSO) |
   | from report | A name was written on the report ("Done By" column, or `01:00 Lt. Afzal`) |
   | ⚠ choose (red) | Two or more people at the site, so you must pick |
   | rotate / random / auto | Filled in by a setting, change it if it's wrong |
   | set by you | You picked it by hand. **Reset** puts it back |

3. **Pick from the dropdown.** It lists the site's officers and Lt. first, then the site's
   dispatcher (dispatchers never count toward the one-person rule, but you can pick them),
   then everyone else on duty. Any row can be changed by hand, including automatic ones.
4. **Copy DAR** or **Save .txt…** unlock once nothing is red.

Other controls:
- **Show only rows that need a pick** hides the automatic rows.
- **Reload** (F5) re-reads the report after you edit it in Excel. Your picks are kept.
- **File ▸ Clear all hand picks** removes every hand pick.

### Settings (only affect sites with 2+ people)

| Setting | Ask me | Rotate | Random | Auto |
|---|---|---|---|---|
| PDA tours | blank, you pick | 1st tour officer, 2nd Lt., … | random person at the site | |
| Visual tours | same as PDA | same as PDA | same as PDA (Shuffle re-rolls) | |
| Locks, escorts, tickets | blank, you pick | | | officer first (old parser rules) |

Settings never overwrite a hand pick. They are remembered in `dar_settings.json`.

### Where picks are saved

Your picks are saved automatically next to the report as `<report name>.dar.json`. Reopening
the same report brings them back. If the folder is read-only, the app says so and keeps working.

## Windows

**Run without packaging:** double-click `run_windows.bat`. The first run sets things up,
after that it opens straight away.

**Package into an .exe** (for PCs without Python):

1. On a Windows PC, install Python from python.org (tick "Add python.exe to PATH").
2. Double-click `build_windows.bat`. It takes a few minutes the first time.
3. The app is in `dist\ShiftReportDAR\`. Copy that **whole folder** anywhere, or to other
   PCs, and run `ShiftReportDAR.exe`. No Python is needed on those PCs.

In the packaged folder:
- `report_config.json` next to the .exe is the one used, so you can still edit the rules.
- `dar_settings.json` (your settings) is saved there too.
- If something goes wrong, details go to `dar_app_error.log` there.
- Windows may show "Windows protected your PC" the first time, because the .exe isn't
  signed. Click **More info ▸ Run anyway**.

The .exe must be built on Windows. PyInstaller can't build a Windows app from a Mac.

## Command line (no window)

```
python shift_report_parser.py report.xlsx               # DAR using rotate/auto for 2+ sites
python shift_report_parser.py report.xlsx --use-saved   # DAR using picks saved by the app
python shift_report_parser.py report.xlsx --sections    # show the sections found on the sheet
```

## Changing the rules: `report_config.json`

| To… | Edit |
|---|---|
| Add a building or street to a campus | `campuses.<Campus>.nicknames` / `keywords` |
| Add a site / staff line, change its label | `personnel.staff_lines` (`site` = it does tours) |
| Change who counts for the one-person rule | `personnel.site_roles` (default officer, supervisor) |
| Roles that are hand-pick only | `personnel.manual_only_roles` (default dispatcher) |
| Where a site with nobody on it falls back to | `site.fallback`, else `personnel.fallback_site` |
| Building with its own order (like 34R Prentiss) | `location_rules` |
| Rotation order | `sections[].rotation` |
| DAR wording | `sections[].template` (a piece is dropped if any of its `{fields}` is empty) |
| Put a new section in the DAR | add an entry to `sections` |
| Default settings | `settings.<category>.default` (app) / `cli_default` (command line) |

How sections are found: every section title on the sheet is a merged cell across the whole
form width. Rows can be inserted anywhere; sections, columns (found by header text) and
the number of tours per row (even past the last column) are all read from the sheet.

## Files

| File | What it is |
|---|---|
| `dar_app.py` | The app window (tkinter) |
| `dar_session.py` | App logic without the window: picks, settings, saving |
| `shift_report_parser.py` | The engine: reads the sheet, builds tasks, writes the DAR |
| `report_config.json` | All report-specific data and wording |
| `tests/` | Automated tests (`pip install pytest`, then `python -m pytest tests`) |
