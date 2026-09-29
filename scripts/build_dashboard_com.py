"""
Adds the fully-interactive Dashboard to Euro_2024_Cross_Dashboard.xlsx using
Excel COM automation (PivotTables, Slicers, native charts, conditional-format
heatmaps) -- these Excel-native objects cannot be created by openpyxl.

Run AFTER build_workbook_data.py has produced the base workbook.
Requires a local installation of Microsoft Excel (uses win32com).
"""
from pathlib import Path
import math
import win32com.client as win32
from win32com.client import dynamic as win32dynamic

ROOT = Path(__file__).resolve().parent.parent
FILE = ROOT / "Euro_2024_Cross_Dashboard.xlsx"

# ---- Excel constant values (hard-coded so we don't depend on gencache) ----
xlDatabase = 1
xlRowField, xlColumnField, xlPageField, xlDataField = 1, 2, 3, 4
xlSum = -4157
xlDescending = 2
xlAutomatic = -4105
xlTop = 1
xlColumnClustered = 51
xlBarClustered = 57
xlPie = 5
xlXYScatter = -4169
xlXYScatterLinesNoMarkers = 75
xlCategory = 1
xlValue = 2
xlLegendPositionBottom = -4107
xlLegendPositionRight = -4152
xlSlicerCrossFilterHideButtonsWithNoData = 4
xlMarkerStyleCircle = 8
xlMarkerStyleX = -4168
xlMarkerStyleNone = -4142
xlCenter = -4108
xlLeft = -4131
xlTickLabelPositionLow = -4134
xlTickLabelPositionNone = -4142


def RGB(r, g, b):
    return r + (g << 8) + (b << 16)


NAVY_RGB = RGB(0x1B, 0x2A, 0x4A)
STEEL_RGB = RGB(0x2E, 0x4A, 0x7A)
GOLD_RGB = RGB(0xC9, 0xA2, 0x27)
GREEN_RGB = RGB(0x2E, 0xA0, 0x59)
RED_RGB = RGB(0xC0, 0x3B, 0x2E)
GRAY_RGB = RGB(0xB5, 0xB9, 0xC2)
LIGHT_RGB = RGB(0xF4, 0xF6, 0xFA)
WHITE_RGB = RGB(0xFF, 0xFF, 0xFF)
INK_RGB = RGB(0x22, 0x22, 0x22)
MUTED_RGB = RGB(0x6B, 0x6B, 0x6B)


def build_pitch_lines_sheet(wb):
    ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
    ws.Name = "PitchLines"

    pts = []

    def line(*coords):
        pts.extend(coords)
        pts.append(None)

    line((0, 0), (0, 80), (120, 80), (120, 0), (0, 0))
    line((60, 0), (60, 80))
    circle = [(60 + 10 * math.cos(t), 40 + 10 * math.sin(t)) for t in
              [i * (2 * math.pi / 36) for i in range(37)]]
    line(*circle)
    line((0, 18), (18, 18), (18, 62), (0, 62))
    line((120, 18), (102, 18), (102, 62), (120, 62))
    line((0, 30), (6, 30), (6, 50), (0, 50))
    line((120, 30), (114, 30), (114, 50), (120, 50))

    ws.Range("A1").Value = "PX"
    ws.Range("B1").Value = "PY"
    row = 2
    for p in pts:
        if p is None:
            continue
        ws.Cells(row, 1).Value = p[0]
        ws.Cells(row, 2).Value = 80 - p[1]  # flipped to match the top=left cross_y_* convention
        row += 1
    ws.Visible = 0
    return ws, row - 1


def style_title_band(ws):
    ws.Range("A1").ColumnWidth = 2.5
    ws.Range("B2:AD3").Merge()
    r = ws.Range("B2")
    r.Value = "EURO 2024 \u2014 CROSS & PASSING INTELLIGENCE DASHBOARD"
    r.Font.Size = 20
    r.Font.Bold = True
    r.Font.Color = WHITE_RGB
    r.Font.Name = "Calibri"
    ws.Range("B2:AD3").Interior.Color = NAVY_RGB
    r.VerticalAlignment = xlCenter
    r.HorizontalAlignment = xlLeft
    r.IndentLevel = 1

    ws.Range("B4:AD4").Merge()
    sub = ws.Range("B4")
    sub.Value = ("Data: StatsBomb open data (via github.com/hudl/open-data) \u00b7 51 matches \u00b7 53,888 passes \u00b7 "
                 "Use the slicers to filter by team, match date, period and side \u2014 every visual recalculates live.")
    sub.Font.Italic = True
    sub.Font.Size = 10
    sub.Font.Color = MUTED_RGB


def section_label(ws, cellref, text, span_cols=6):
    rng = ws.Range(cellref)
    col = rng.Column
    row = rng.Row
    end_col = col + span_cols - 1
    merge_range = ws.Range(ws.Cells(row, col), ws.Cells(row, end_col))
    merge_range.Merge()
    merge_range.Value = text
    merge_range.Font.Bold = True
    merge_range.Font.Size = 11
    merge_range.Font.Color = WHITE_RGB
    merge_range.Interior.Color = STEEL_RGB
    merge_range.HorizontalAlignment = xlLeft
    merge_range.IndentLevel = 1
    merge_range.RowHeight = 20


def add_table_slicer(wb, ws_dest, lo, field_name, caption, left, top, width=190, height=140):
    sc = wb.SlicerCaches.Add2(lo, field_name)
    name = f"Slicer_{field_name}_{ws_dest.Name}"
    sc.Slicers.Add(SlicerDestination=ws_dest, Name=name, Caption=caption, Top=top, Left=left, Width=width, Height=height)
    sc.CrossFilterType = xlSlicerCrossFilterHideButtonsWithNoData
    return sc


def add_pivot_slicer(wb, ws_dest, pivot, field_name, caption, left, top, width=190, height=140):
    sc = wb.SlicerCaches.Add2(pivot, field_name)
    name = f"Slicer_{field_name}_pivot_{ws_dest.Name}"
    sc.Slicers.Add(SlicerDestination=ws_dest, Name=name, Caption=caption, Top=top, Left=left, Width=width, Height=height)
    sc.CrossFilterType = xlSlicerCrossFilterHideButtonsWithNoData
    return sc


def kpi_card(ws, top_left_cell, formula, label, number_format="0", fill=WHITE_RGB, value_color=NAVY_RGB):
    r = ws.Range(top_left_cell)
    row, col = r.Row, r.Column
    value_rng = ws.Range(ws.Cells(row, col), ws.Cells(row + 1, col + 2))
    value_rng.Merge()
    value_rng.Formula = formula
    value_rng.NumberFormat = number_format
    value_rng.Font.Size = 26
    value_rng.Font.Bold = True
    value_rng.Font.Color = value_color
    value_rng.HorizontalAlignment = xlCenter
    value_rng.VerticalAlignment = xlCenter
    value_rng.Interior.Color = fill
    value_rng.Borders.Weight = 2
    value_rng.Borders.Color = RGB(0xDD, 0xDD, 0xE2)

    label_rng = ws.Range(ws.Cells(row + 2, col), ws.Cells(row + 2, col + 2))
    label_rng.Merge()
    label_rng.Value = label.upper()
    label_rng.Font.Size = 9
    label_rng.Font.Bold = True
    label_rng.Font.Color = MUTED_RGB
    label_rng.HorizontalAlignment = xlCenter


def apply_color_scale(rng, low=WHITE_RGB, mid=RGB(0xFF, 0xD9, 0x6B), high=RGB(0xC0, 0x3B, 0x2E)):
    rng.FormatConditions.Delete()
    fc = rng.FormatConditions.AddColorScale(3)
    fc.ColorScaleCriteria(1).FormatColor.Color = low
    fc.ColorScaleCriteria(2).FormatColor.Color = mid
    fc.ColorScaleCriteria(3).FormatColor.Color = high
    rng.NumberFormat = "0"
    rng.Font.Size = 10
    rng.HorizontalAlignment = xlCenter


def build_zone_pivot(pc, ws, dest_cell, table_name, value_field, value_caption, title):
    r = ws.Range(dest_cell)
    ws.Cells(r.Row - 1, r.Column).Value = title
    ws.Cells(r.Row - 1, r.Column).Font.Bold = True
    ws.Cells(r.Row - 1, r.Column).Font.Size = 12
    ws.Cells(r.Row - 1, r.Column).Font.Color = NAVY_RGB

    pt = pc.CreatePivotTable(TableDestination=r, TableName=table_name)
    pt.PivotFields("y_band").Orientation = xlRowField
    pt.PivotFields("x_band").Orientation = xlColumnField
    pt.AddDataField(pt.PivotFields(value_field), value_caption, xlSum)
    pt.RowGrand = False
    pt.ColumnGrand = False
    pt.TableStyle2 = "PivotStyleLight16"
    pt.DataFields.Item(1).NumberFormat = "0"
    return pt


def build_flat_pivot(pc, ws, dest_cell, table_name, row_field, data_fields, title, col_field=None):
    r = ws.Range(dest_cell)
    ws.Cells(r.Row - 1, r.Column).Value = title
    ws.Cells(r.Row - 1, r.Column).Font.Bold = True
    ws.Cells(r.Row - 1, r.Column).Font.Size = 12
    ws.Cells(r.Row - 1, r.Column).Font.Color = NAVY_RGB

    pt = pc.CreatePivotTable(TableDestination=r, TableName=table_name)
    pt.PivotFields(row_field).Orientation = xlRowField
    if col_field:
        pt.PivotFields(col_field).Orientation = xlColumnField
    for field, caption in data_fields:
        pt.AddDataField(pt.PivotFields(field), caption, xlSum)
    pt.RowGrand = True
    pt.ColumnGrand = False if col_field else True
    pt.TableStyle2 = "PivotStyleLight16"
    for i in range(1, len(data_fields) + 1):
        pt.DataFields.Item(i).NumberFormat = "0"
    return pt


def add_bar_chart(ws, pt, left, top, width, height, title, chart_type=xlColumnClustered, max_rows=None):
    rng = pt.TableRange1
    if max_rows:
        n_cols = rng.Columns.Count
        rng = ws.Range(rng.Cells(1, 1), rng.Cells(min(max_rows + 1, rng.Rows.Count), n_cols))
    shp = ws.Shapes.AddChart2(-1, chart_type, left, top, width, height)
    chart = shp.Chart
    chart.SetSourceData(rng)
    chart.HasTitle = True
    chart.ChartTitle.Text = title
    chart.ChartTitle.Font.Size = 12
    chart.ChartTitle.Font.Bold = True
    chart.HasLegend = True
    chart.Legend.Position = xlLegendPositionBottom
    chart.ChartArea.Format.Fill.ForeColor.RGB = WHITE_RGB
    return chart


def main():
    print("Starting Excel...")
    excel = win32dynamic.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False

    wb = excel.Workbooks.Open(str(FILE))
    ws_passes = wb.Sheets("Passes")
    lo = ws_passes.ListObjects("tbl_Passes")
    print("Table rows:", lo.ListRows.Count)

    # ---------------- Dashboard sheet shell ----------------
    ws = wb.Sheets.Add(Before=ws_passes)
    ws.Name = "Dashboard"
    ws.Cells.Font.Name = "Calibri"
    style_title_band(ws)

    print("Building pitch line helper sheet...")
    ws_pitch, pitch_n = build_pitch_lines_sheet(wb)

    # ---------------- Group 1: table-driven filters (pitch map + KPIs) --------
    section_label(ws, "B6", "PITCH MAP & KEY PERFORMANCE INDICATORS  \u2014  FILTERS", span_cols=20)
    g1_top = ws.Range("B7").Top
    g1_left0 = ws.Range("B7").Left
    g1_fields = [
        ("team", "Team"), ("match_label", "Match (date order)"),
        ("half_label", "Period"), ("side", "Side"), ("cross", "Cross?"),
    ]
    for i, (field, caption) in enumerate(g1_fields):
        add_table_slicer(wb, ws, lo, field, caption, g1_left0 + i * 200, g1_top, 190, 145)
    print("Group 1 slicers added.")

    # KPI cards
    kpi_row_cell = "B18"
    kpi_top = ws.Range(kpi_row_cell).Top
    section_label(ws, "B17", "KEY PERFORMANCE INDICATORS (reflect current filter selection)", span_cols=20)
    kpi_specs = [
        ("B19", "=SUBTOTAL(103,tbl_Passes[match_id])", "Passes (Filtered)", "#,##0", NAVY_RGB),
        ("F19", "=SUBTOTAL(109,tbl_Passes[cross_flag])", "Crosses (Filtered)", "#,##0", NAVY_RGB),
        ("J19", "=IFERROR(SUBTOTAL(109,tbl_Passes[cross_completed_flag])/SUBTOTAL(109,tbl_Passes[cross_flag]),0)",
         "Cross Completion %", "0.0%", GREEN_RGB),
        ("N19", "=IFERROR(SUBTOTAL(109,tbl_Passes[cross_box_entry_flag])/SUBTOTAL(109,tbl_Passes[cross_flag]),0)",
         "Crosses Targeting Box", "0.0%", GOLD_RGB),
        ("R19", "=SUBTOTAL(109,tbl_Passes[cross_dangerous_flag])", "Dangerous Crosses", "#,##0", RED_RGB),
    ]
    for cell, formula, label, fmt, color in kpi_specs:
        kpi_card(ws, cell, formula, label, fmt, WHITE_RGB, color)
    print("KPI cards added.")

    # Pitch scatter chart
    ws.Range("B23").Value = "Cross Origin Pitch Map \u2014 completed vs incomplete (goal on the right, left side of attack shown at top, right side at bottom, matching the zone-grid convention)"
    ws.Range("B23").Font.Bold = True
    ws.Range("B23").Font.Size = 12
    ws.Range("B23").Font.Color = NAVY_RGB

    chart_left = ws.Range("B24").Left
    chart_top = ws.Range("B24").Top
    shp = ws.Shapes.AddChart2(-1, xlXYScatter, chart_left, chart_top, 560, 373)
    chart = shp.Chart
    s1 = chart.SeriesCollection().NewSeries()
    s1.XValues = lo.ListColumns("cross_x_plot").DataBodyRange
    s1.Values = lo.ListColumns("cross_y_completed").DataBodyRange
    s1.Name = "Completed cross"
    s1.MarkerStyle = xlMarkerStyleCircle
    s1.MarkerSize = 6
    s1.MarkerForegroundColor = GREEN_RGB
    s1.MarkerBackgroundColor = GREEN_RGB
    s1.Format.Line.Visible = False

    s2 = chart.SeriesCollection().NewSeries()
    s2.XValues = lo.ListColumns("cross_x_plot").DataBodyRange
    s2.Values = lo.ListColumns("cross_y_incomplete").DataBodyRange
    s2.Name = "Incomplete cross"
    s2.MarkerStyle = xlMarkerStyleX
    s2.MarkerSize = 6
    s2.MarkerForegroundColor = RED_RGB
    s2.MarkerBackgroundColor = RED_RGB
    s2.Format.Line.Visible = False

    s3 = chart.SeriesCollection().NewSeries()
    s3.XValues = ws_pitch.Range(f"A2:A{pitch_n}")
    s3.Values = ws_pitch.Range(f"B2:B{pitch_n}")
    s3.Name = "Pitch"
    s3.ChartType = xlXYScatterLinesNoMarkers
    s3.MarkerStyle = xlMarkerStyleNone
    s3.Format.Line.ForeColor.RGB = GRAY_RGB
    s3.Format.Line.Weight = 1.25

    xaxis = chart.Axes(xlCategory)
    xaxis.MinimumScale = 0
    xaxis.MaximumScale = 120
    xaxis.MajorUnit = 20
    xaxis.TickLabelPosition = xlTickLabelPositionLow
    yaxis = chart.Axes(xlValue)
    yaxis.MinimumScale = 0
    yaxis.MaximumScale = 80
    yaxis.MajorUnit = 16

    chart.HasTitle = False
    chart.HasLegend = True
    chart.Legend.Position = xlLegendPositionBottom
    try:
        chart.Legend.LegendEntries(3).Delete()
    except Exception as e:
        print("legend entry delete skipped:", e)
    chart.ChartArea.Format.Fill.ForeColor.RGB = WHITE_RGB
    chart.PlotArea.Format.Fill.ForeColor.RGB = RGB(0xEE, 0xF3, 0xEC)
    print("Pitch scatter chart added.")

    # ---------------- Group 2: pivot-driven zone heatmaps & breakdowns --------
    ws2g = wb.Sheets.Add(After=ws)
    ws2g.Name = "Zone Analytics"
    ws2g.Cells.Font.Name = "Calibri"
    style_title_band(ws2g)

    pc = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData="tbl_Passes")

    section_label(ws2g, "B6", "ZONE HEATMAP & BREAKDOWN ANALYTICS  \u2014  FILTERS", span_cols=20)
    g2_top = ws2g.Range("B7").Top
    g2_left0 = ws2g.Range("B7").Left
    g2_fields = [("team", "Team"), ("match_label", "Match (date order)"), ("half_label", "Period"), ("side", "Side")]

    pt_zone_all = build_zone_pivot(pc, ws2g, "B19", "pt_ZoneAll", "cross_flag", "Total Crosses",
                                    "Zone Heatmap \u2014 All Crosses")
    apply_color_scale(pt_zone_all.DataBodyRange)
    print("Zone heatmap (all) pivot added.")

    pt_zone_comp = build_zone_pivot(pc, ws2g, "N19", "pt_ZoneCompleted", "cross_completed_flag", "Completed Crosses",
                                     "Zone Heatmap \u2014 Completed Crosses Only")
    apply_color_scale(pt_zone_comp.DataBodyRange)
    print("Zone heatmap (completed) pivot added.")

    pt_delivery = build_flat_pivot(pc, ws2g, "B31", "pt_Delivery", "cross_channel",
                                    [("cross_flag", "Total Crosses")], "Delivery Profile \u2014 Channel \u00d7 Depth",
                                    col_field="cross_depth")
    apply_color_scale(pt_delivery.DataBodyRange, mid=RGB(0x8F, 0xBF, 0xEA), high=RGB(0x00, 0x71, 0xE3))
    print("Delivery profile pivot added.")

    pt_half = build_flat_pivot(pc, ws2g, "N31", "pt_Half",
                                "half_label", [("cross_flag", "Total Crosses"), ("cross_completed_flag", "Completed Crosses")],
                                "Crosses by Match Period")
    add_bar_chart(ws2g, pt_half, ws2g.Range("U31").Left, ws2g.Range("U31").Top, 380, 240,
                  "Crosses by Match Period")
    print("Half pivot + chart added.")

    pt_side = build_flat_pivot(pc, ws2g, "B44", "pt_Side",
                                "side", [("cross_flag", "Total Crosses"), ("cross_completed_flag", "Completed Crosses")],
                                "Crosses by Side")
    add_bar_chart(ws2g, pt_side, ws2g.Range("F44").Left, ws2g.Range("F44").Top, 380, 240,
                  "Crosses by Side")
    print("Side pivot + chart added.")

    pt_player = build_flat_pivot(pc, ws2g, "N44", "pt_TopPlayers",
                                  "player", [("cross_flag", "Total Crosses"), ("cross_completed_flag", "Completed Crosses")],
                                  "Top 10 Crossers")
    pt_player.RowGrand = False
    pt_player.PivotFields("player").AutoSort(xlDescending, "Total Crosses")
    add_bar_chart(ws2g, pt_player, ws2g.Range("R44").Left, ws2g.Range("R44").Top, 460, 300,
                  "Top 10 Crossers", chart_type=xlBarClustered, max_rows=10)
    print("Top players pivot + chart added.")

    for i, (field, caption) in enumerate(g2_fields):
        if i == 0:
            sc2 = add_pivot_slicer(wb, ws2g, pt_zone_all, field, caption, g2_left0 + i * 210, g2_top, 200, 145)
        else:
            sc2 = wb.SlicerCaches.Add2(pt_zone_all, field)
            sc2.Slicers.Add(SlicerDestination=ws2g, Name=f"Slicer_{field}_pivot_{ws2g.Name}", Caption=caption,
                             Top=g2_top, Left=g2_left0 + i * 210, Width=200, Height=145)
            sc2.CrossFilterType = xlSlicerCrossFilterHideButtonsWithNoData
        for other_pt in (pt_zone_comp, pt_delivery, pt_half, pt_side, pt_player):
            try:
                sc2.PivotTables.AddPivotTable(other_pt)
            except Exception as e:
                print(f"connect slicer {field} to {other_pt.Name} failed:", e)
    print("Group 2 slicers added + connected across all pivots.")

    ws2g.Range("A1").Select()
    ws.Activate()
    ws.Range("A1").Select()
    wb.Save()
    print("Phase 2+3 (full dashboard) saved OK.")
    return excel, wb, ws, ws_passes, lo, ws_pitch, pitch_n


if __name__ == "__main__":
    excel, wb, ws, ws_passes, lo, ws_pitch, pitch_n = main()
    wb.Close(SaveChanges=True)
    excel.Quit()
    print("Done (phase 2).")
