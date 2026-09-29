# Calcio IQ — Euro 2024 Cross & Passing Intelligence

A coach-facing analytics package built on StatsBomb's Euro 2024 open event data,
created for the ELVTR Soccer Performance Analytics guest lecture (Jordyn Kaplan).

## Contents

- **`Euro_2024_Analytics_Report.html`** — self-contained, interactive HTML report.
  Open it directly in a browser, or view it live via GitHub Pages (see below).
  Slicer-style tiles for Team / Match / Period / Side cross-filter every chart,
  KPI, and the League Overview table; includes a per-team coach's tactical summary.
- **`Euro_2024_Cross_Dashboard.xlsx`** — fully interactive Excel workbook with
  native PivotTables, Slicers, conditional-formatting zone heatmaps, and a
  pitch scatter chart, all driven by Team/Match/Period/Side slicers.
- **`euro_2024_passes.xlsx`** / **`teams.xlsx`** — raw source data.
- **`Cross_Map_Instructions_Excel_Sheets.docx`** — original assignment brief.
- **`scripts/`** — Python build scripts that regenerate both deliverables from
  the raw data (`data_prep.py`, `build_workbook_data.py`, `build_dashboard_com.py`,
  `build_html_report.py`).

## Live report

Once GitHub Pages is enabled for this repo (Settings → Pages → Deploy from
branch `main` / root), the report is viewable at:

```
https://rolomagnus.github.io/Calcio-IQ/Euro_2024_Analytics_Report.html
```

## Rebuilding the deliverables

```powershell
cd scripts
python build_workbook_data.py   # writes the base Excel workbook (openpyxl)
python build_dashboard_com.py   # adds PivotTables/Slicers/charts (requires Excel + Windows)
python build_html_report.py     # regenerates the HTML report
```

## Data credit

Data: StatsBomb open data via [github.com/hudl/open-data](https://github.com/hudl/open-data),
free for non-commercial use. This analysis is based on pass/cross event data only —
it does not include final match scores or whether a cross led to a goal.
