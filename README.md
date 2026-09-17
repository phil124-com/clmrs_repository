# CLMRS Data Analyzer

A reusable desktop tool for analysing Child Labour Monitoring and Remediation
System (CLMRS) data collected through Excel-based forms — built for the
Butter (LS Country) programme first, but reusable for any programme with the
same or similar structure via the Programme Configuration system.

## What it does

Upload your four datasets — **Farmer List (Annex)**, **Household**,
**Inspection**, **Follow Up** — click **Run Complete Analysis**, then
**Export Client Report** for a formatted Excel workbook covering:

- Farmer profile analysis (unique farmers, with/without children)
- Children profile analysis (gender, school attendance)
- Farmer List / Annex matching (present/absent)
- Inspection analysis, and Household-vs-Inspection coverage comparison
- Follow Up case analysis with calendar-month-accurate Case Identified Date
- Quarter Analysis and Enumerator Analysis
- Data Quality exceptions

## Getting the Windows .exe — no Python, no command line

You don't need to install anything to **use** the finished app. To **build**
it into a `.exe` the first time (or after you change anything), you need a
free GitHub account:

1. **Create a GitHub repository** and upload this whole folder to it
   (github.com → New repository → "uploading an existing file", drag every
   file/folder in, then click Commit).
2. Go to the **Actions** tab of your repository. A workflow called
   "Build Windows Executable" runs automatically (it also re-runs any time
   you push a change). Wait for the green checkmark (a few minutes).
3. Click the finished run → under **Artifacts**, download
   `CLMRS_Data_Analyzer-windows`. Unzip it — inside is
   `CLMRS_Data_Analyzer.exe`. Copy that to your desktop and double-click it
   to run. You can right-click → "Send to → Desktop (create shortcut)" for a
   permanent shortcut.

No Python, pip, or terminal is needed on your Windows laptop at any point —
GitHub's own servers do the building.

## Running it locally for testing (optional, for development only)

```
pip install -r requirements.txt
python main.py
```

## Adding a new programme (e.g. MARS, Lindt, Ferrero)

Programme configurations live in `%APPDATA%\CLMRS_Analyzer\programmes\*.json`
once the app has run once. Copy `Butter.json` to a new file named after the
programme and adjust `survey_header_map` / `farmerlist_header_map` if that
programme's export uses different column headers — the app will list it in
the Programme dropdown automatically. If the new programme's file structure
matches Butter's exactly, no changes are needed at all.

## Project structure

```
main.py                  entry point (what PyInstaller packages)
gui/app.py                Tkinter desktop interface
engine/                    data loading, mapping, analysis, export
config/programmes.py       programme configuration system
build/clmrs_analyzer.spec  PyInstaller build spec
.github/workflows/         GitHub Actions build (produces the .exe)
```
