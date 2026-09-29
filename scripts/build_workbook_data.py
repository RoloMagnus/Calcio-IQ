"""
Builds the base workbook (data only) using openpyxl:
  - Passes: full enriched dataset as a native Excel Table (ListObject)
  - League Overview: team-level KPI leaderboard + top crossers table
  - Read Me: instructions / credit

PivotTables, Slicers, native charts and conditional-formatting heatmaps are
added afterwards by build_dashboard_com.py (requires Excel via COM, since
openpyxl cannot create PivotTables/Slicers).
"""
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference

from data_prep import load_enriched_passes, team_summary, top_players

ROOT = Path(__file__).resolve().parent.parent
OUT_FILE = ROOT / "Euro_2024_Cross_Dashboard.xlsx"

NAVY = "1B2A4A"
GOLD = "C9A227"
LIGHT = "F4F6FA"


def style_header(ws, row, ncols, fill=NAVY, font_color="FFFFFF"):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(bold=True, color=font_color, size=11)
        cell.fill = PatternFill("solid", fgColor=fill)
        cell.alignment = Alignment(horizontal="center", vertical="center")


def write_df_as_table(ws, df: pd.DataFrame, table_name: str, start_row=1, start_col=1, style="TableStyleMedium9"):
    rows = list(dataframe_to_rows(df, index=False, header=True))
    for r_off, row in enumerate(rows):
        for c_off, value in enumerate(row):
            if isinstance(value, float) and pd.isna(value):
                value = None
            ws.cell(row=start_row + r_off, column=start_col + c_off, value=value)

    n_rows = len(rows)
    n_cols = len(df.columns)
    first_col_letter = get_column_letter(start_col)
    last_col_letter = get_column_letter(start_col + n_cols - 1)
    ref = f"{first_col_letter}{start_row}:{last_col_letter}{start_row + n_rows - 1}"

    tab = Table(displayName=table_name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(
        name=style, showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False
    )
    ws.add_table(tab)
    return ref


def autofit(ws, df, start_col=1, max_width=32):
    for i, col in enumerate(df.columns):
        width = min(max(len(str(col)), df[col].astype(str).str.len().quantile(0.95) if len(df) else 10) + 3, max_width)
        ws.column_dimensions[get_column_letter(start_col + i)].width = width


def build():
    print("Loading + enriching data...")
    df = load_enriched_passes()
    df_out = df.drop(columns=["cross_y_completed", "cross_y_incomplete", "cross_x_plot"]).copy()
    # keep plotting helper columns at the end (used by the pitch scatter chart)
    df_out["cross_x_plot"] = df["cross_x_plot"]
    df_out["cross_y_completed"] = df["cross_y_completed"]
    df_out["cross_y_incomplete"] = df["cross_y_incomplete"]

    ts = team_summary(df)
    tp = top_players(df)

    wb = Workbook()

    # ---------------- Read Me ----------------
    ws = wb.active
    ws.title = "Read Me"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 110
    lines = [
        ("Euro 2024 — Cross & Passing Intelligence Dashboard", 18, True),
        ("", 11, False),
        ("Data source: StatsBomb open data via github.com/hudl/open-data (free for non-commercial use).", 11, False),
        ("Assignment: ELVTR Soccer Performance Analytics — Guest Lecture, Jordyn Kaplan.", 11, False),
        ("", 11, False),
        ("How to use this workbook", 13, True),
        ("1. Go to the Dashboard tab. Use the slicers at the top to filter by Team, Match (in chronological", 11, False),
        ("   order), Period/Half, Side and whether the touch was a Cross.", 11, False),
        ("2. The pitch map, KPI cards, zone heatmaps, and breakdown charts all recalculate instantly.", 11, False),
        ("3. League Overview shows every team side-by-side (unaffected by the Team slicer) so you can", 11, False),
        ("   benchmark one team's crossing output against the rest of the tournament.", 11, False),
        ("4. Passes is the single source table (53,888 rows) all visuals are built from — every additional", 11, False),
        ("   support column (side, completed, zone, half_label, cross_depth, cross_channel, box_entry,", 11, False),
        ("   dangerous_cross, match_label, x_band/y_band) is documented in the Data Dictionary below.", 11, False),
        ("", 11, False),
        ("Data Dictionary — support columns added on top of the raw StatsBomb export", 13, True),
        ("side               Left/Right half of the pitch the pass started in (y < 40 = Left).", 11, False),
        ("completed          Yes/No — a blank StatsBomb outcome means the pass was completed.", 11, False),
        ("zone               6x5 pitch grid cell the pass started in, e.g. '4-3' (x-band, y-band).", 11, False),
        ("x_band / y_band    Numeric bin starts (0,20,40,60,80,100 / 0,16,32,48,64) used to drive the", 11, False),
        ("                   zone heatmap pivot grids so columns/rows sort left-to-right, top-to-bottom.", 11, False),
        ("match_label        'YYYY-MM-DD  Home vs Away' — sorts chronologically as-is, used for the", 11, False),
        ("                   Match slicer so matches are always filtered in true date order.", 11, False),
        ("half_label         1st Half / 2nd Half / Extra Time 1 / Extra Time 2 (from period 1-4).", 11, False),
        ("pitch_third        Defensive / Middle / Attacking third the pass started in.", 11, False),
        ("cross_depth        Where a cross was struck from: Deep / Advanced / By-line.", 11, False),
        ("cross_channel      Where a cross was struck from: Wide Channel / Half-Space / Central.", 11, False),
        ("box_entry          Yes/No — does the pass end point (end_x, end_y) land inside the penalty box", 11, False),
        ("                   (x >= 102, 18 <= y <= 62)? Measures targeting, independent of completion.", 11, False),
        ("dangerous_cross    Yes/No — completed AND targeted the box: the cleanest 'quality delivery' flag.", 11, False),
        ("cross_x_plot,      Helper columns feeding the pitch scatter chart (NA when not a cross, or when", 11, False),
        ("cross_y_completed, the outcome doesn't match, so the chart only plots the right series).", 11, False),
        ("cross_y_incomplete", 11, False),
        ("", 11, False),
        ("Coaching questions this dashboard is built to answer", 13, True),
        ("- How much does this team cross, and from which side / channel / depth?", 11, False),
        ("- How often does a cross actually reach the box, and how often does it stay alive (complete)?", 11, False),
        ("- Is delivery concentrated in one player the opposition can scout and stop?", 11, False),
        ("- Does the pattern change first half vs second half (fatigue, chasing the game, extra time)?", 11, False),
        ("- How does this team compare to the rest of the tournament field (League Overview)?", 11, False),
    ]
    r = 1
    for text, size, bold in lines:
        cell = ws.cell(row=r, column=2, value=text)
        cell.font = Font(size=size, bold=bold, color=NAVY if bold else "222222")
        r += 1
    ws.freeze_panes = "A2"

    # ---------------- Passes (data table) ----------------
    ws2 = wb.create_sheet("Passes")
    print(f"Writing {len(df_out):,} rows to Passes...")
    ref = write_df_as_table(ws2, df_out, "tbl_Passes", style="TableStyleMedium2")
    ws2.freeze_panes = "A2"
    style_header(ws2, 1, len(df_out.columns))
    date_col_idx = list(df_out.columns).index("match_date") + 1
    ws2.column_dimensions[get_column_letter(date_col_idx)].number_format = "yyyy-mm-dd"
    widths = {
        "match_id": 11, "match_date": 12, "home_team": 14, "away_team": 14, "team": 14, "opponent": 14,
        "player": 24, "period": 8, "minute": 8, "x": 7, "y": 7, "end_x": 8, "end_y": 8, "cross": 7,
        "outcome": 16, "pass_height": 13, "body_part": 12, "side": 8, "completed": 10, "zone": 8,
        "match_label": 34, "half_label": 14, "pitch_third": 15, "x_band": 8, "y_band": 8,
        "cross_depth": 12, "cross_channel": 13, "box_entry": 10, "dangerous_cross": 14,
        "matches_played": 10, "cross_x_plot": 12, "cross_y_completed": 16, "cross_y_incomplete": 16,
    }
    for i, col in enumerate(df_out.columns):
        ws2.column_dimensions[get_column_letter(i + 1)].width = widths.get(col, 12)
    print("Passes table ref:", ref)

    # ---------------- League Overview ----------------
    ws3 = wb.create_sheet("League Overview")
    ws3.sheet_view.showGridLines = False
    title = ws3.cell(row=1, column=1, value="Euro 2024 — League-Wide Crossing Leaderboard")
    title.font = Font(size=16, bold=True, color=NAVY)
    sub = ws3.cell(row=2, column=1, value="Static, tournament-wide comparison (not affected by the Dashboard slicers).")
    sub.font = Font(size=10, italic=True, color="666666")

    ts_display = ts.rename(columns={
        "team": "Team", "matches_played": "Matches Played", "total_passes": "Total Passes",
        "total_crosses": "Total Crosses", "crosses_per_match": "Crosses / Match",
        "completed_crosses": "Completed Crosses", "completion_pct": "Completion %",
        "left_side": "Left Side", "right_side": "Right Side",
        "box_entries": "Box Entries", "box_entry_pct": "Box Entry %", "dangerous": "Dangerous Crosses",
    })
    # ranked by volume of passes (most passes first), not by crossing output
    ts_display = ts_display.sort_values("Total Passes", ascending=False).reset_index(drop=True)
    col_order = ["Team", "Matches Played", "Total Passes", "Total Crosses", "Crosses / Match",
                 "Completed Crosses", "Completion %", "Left Side", "Right Side",
                 "Box Entries", "Box Entry %", "Dangerous Crosses"]
    ts_display = ts_display[col_order]
    ref2 = write_df_as_table(ws3, ts_display, "tbl_LeagueOverview", start_row=4, style="TableStyleMedium9")
    style_header(ws3, 4, len(col_order))
    autofit(ws3, ts_display)

    # top crossers table alongside
    tp_display = tp.head(20).rename(columns={
        "team": "Team", "player": "Player", "total_crosses": "Total Crosses",
        "completed_crosses": "Completed", "completion_pct": "Completion %",
    })
    start_col_tp = len(col_order) + 3
    hdr = ws3.cell(row=4, column=start_col_tp, value="Top 20 Crossers — Euro 2024 (min. 3 crosses)")
    hdr.font = Font(size=12, bold=True, color=NAVY)
    write_df_as_table(ws3, tp_display, "tbl_TopCrossers", start_row=6, start_col=start_col_tp, style="TableStyleMedium9")
    style_header(ws3, 6, len(tp_display.columns), fill=NAVY)
    for i, col in enumerate(tp_display.columns):
        ws3.column_dimensions[get_column_letter(start_col_tp + i)].width = 16 if col != "Player" else 26

    # simple bar chart: crosses per match ranking
    chart = BarChart()
    chart.title = "Crosses per Match — All 24 Teams"
    chart.y_axis.title = "Crosses / Match"
    chart.style = 10
    n = len(ts_display)
    data = Reference(ws3, min_col=5, min_row=4, max_row=4 + n)
    cats = Reference(ws3, min_col=1, min_row=5, max_row=4 + n)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.width = 32
    chart.height = 14
    ws3.add_chart(chart, f"A{6 + n + 2}")

    wb.save(OUT_FILE)
    print("Saved base workbook:", OUT_FILE)


if __name__ == "__main__":
    build()
