"""
Builds a clean, Apple-style, self-contained interactive HTML analytics report
covering every metric surfaced in the Excel dashboard: tournament KPIs, the
league-wide crossing leaderboard, zone heatmaps, delivery profile, half/side
splits, top crossers, and a per-team pitch map -- all driven by Excel-style
slicer tiles (Team / Match / Period / Side) that filter every chart, KPI and
table row on the page, mirroring the workbook's slicers exactly.
"""
import json
from pathlib import Path
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio

from data_prep import load_enriched_passes, team_summary

ROOT = Path(__file__).resolve().parent.parent
OUT_FILE = ROOT / "Euro_2024_Analytics_Report.html"

NAVY = "#1B2A4A"
INK = "#1D1D1F"
MUTED = "#6E6E73"
BLUE = "#0071E3"
GREEN = "#2EA059"
RED = "#C0392B"
GOLD = "#C9A227"
GRAY = "#D2D2D7"
CARD = "#FFFFFF"

FONT = "-apple-system, BlinkMacSystemFont, 'SF Pro Display', 'Segoe UI', Helvetica, Arial, sans-serif"
HALF_ORDER = ["1st Half", "2nd Half", "Extra Time 1", "Extra Time 2"]
# Y bands listed BOTTOM -> TOP as Plotly's default heatmap draws the first
# array entry at the bottom and the last at the top. Ordering the bands this
# way (64-80 first, 0-16 last) puts the 0-16 ("Left") band at the TOP of the
# heatmap, matching the zone-grid convention (top rows = left side of attack)
# and the pitch map below (which also shows the left side at the top).
ZONE_Y = [64, 48, 32, 16, 0]
ZONE_X = [0, 20, 40, 60, 80, 100]
CHANNEL_ORDER = ["Wide Channel", "Half-Space", "Central"]
DEPTH_ORDER = ["Deep", "Advanced", "By-line"]

pio.templates.default = "plotly_white"


def fig_to_div(fig, div_id):
    fig.update_layout(
        font=dict(family=FONT, color=INK, size=13),
        margin=dict(l=50, r=30, t=60, b=50),
        paper_bgcolor=CARD, plot_bgcolor=CARD,
        title_font=dict(size=16, color=NAVY),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    return pio.to_html(fig, full_html=False, include_plotlyjs=False, div_id=div_id, config={"displaylogo": False})


def pitch_shapes():
    return [
        dict(type="rect", x0=0, y0=0, x1=120, y1=80, line=dict(color="#C9CDD3", width=2)),
        dict(type="line", x0=60, y0=0, x1=60, y1=80, line=dict(color="#C9CDD3", width=1.5)),
        dict(type="circle", x0=50, y0=30, x1=70, y1=50, line=dict(color="#C9CDD3", width=1.5)),
        dict(type="rect", x0=0, y0=18, x1=18, y1=62, line=dict(color="#C9CDD3", width=1.5)),
        dict(type="rect", x0=102, y0=18, x1=120, y1=62, line=dict(color="#C9CDD3", width=1.5)),
        dict(type="rect", x0=0, y0=30, x1=6, y1=50, line=dict(color="#C9CDD3", width=1.5)),
        dict(type="rect", x0=114, y0=30, x1=120, y1=50, line=dict(color="#C9CDD3", width=1.5)),
    ]


def zone_matrix(sub):
    if len(sub) == 0:
        return [[0] * len(ZONE_X) for _ in ZONE_Y]
    z = sub.pivot_table(index="y_band", columns="x_band", values="cross_flag", aggfunc="sum", fill_value=0)
    z = z.reindex(index=ZONE_Y, columns=ZONE_X, fill_value=0)
    return z.astype(int).values.tolist()


def delivery_matrix(sub):
    if len(sub) == 0:
        return [[0] * len(DEPTH_ORDER) for _ in CHANNEL_ORDER]
    dp = sub.pivot_table(index="cross_channel", columns="cross_depth", values="cross_flag", aggfunc="sum", fill_value=0)
    dp = dp.reindex(index=CHANNEL_ORDER, columns=DEPTH_ORDER, fill_value=0)
    return dp.astype(int).values.tolist()


def half_data(sub):
    agg = sub.groupby("half_label").agg(
        total=("cross_flag", "sum"), completed=("cross_completed_flag", "sum")
    ).reindex(HALF_ORDER).fillna(0)
    return agg["total"].astype(int).tolist(), agg["completed"].astype(int).tolist()


def players_data(sub, n, show_team_suffix):
    if len(sub) == 0:
        return [], [], []
    agg = sub.groupby(["player", "team"]).agg(
        total=("cross_flag", "sum"), completed=("cross_completed_flag", "sum")
    ).reset_index()
    agg = agg[agg["total"] > 0].sort_values("total", ascending=False).head(n)
    agg = agg.sort_values("total", ascending=True)
    if show_team_suffix:
        labels = (agg["player"] + " (" + agg["team"] + ")").tolist()
    else:
        labels = agg["player"].tolist()
    return labels, agg["total"].astype(int).tolist(), agg["completed"].astype(int).tolist()


CAVEAT_HTML = (
    '<p class="caveat">\u26a0\ufe0f <b>Data caveat:</b> this analysis is built from pass/cross event data only. '
    'It does not include final match scores, results, or whether any cross led directly to a shot or goal \u2014 '
    'so end-product (finishing quality) cannot be assessed from this dataset alone. Treat "dangerous crosses" '
    '(completed AND reaching the box) as the closest proxy for delivery quality, not a confirmed goal contribution.</p>'
)


def build_coach_summaries(df, all_cr, ts):
    """Pre-computed per-team tactical paragraphs, swapped in by the Team slicer,
    answering: how much do they cross and from which side, where from and how
    often does it work, and what should defenders be told."""
    summaries = {}
    total_matches = int(df["match_id"].nunique())
    total_crosses = int(len(all_cr))
    league_left_pct = round((all_cr["side"] == "Left").mean() * 100, 1)
    league_completion = round((all_cr["completed"] == "Yes").mean() * 100, 1)
    top_row = ts.sort_values("crosses_per_match", ascending=False).iloc[0]

    summaries["ALL"] = (
        f"<p>Across all {total_matches} Euro 2024 matches, the field combined for {total_crosses:,} crosses "
        f"({total_crosses / total_matches:.1f} per match), completing {league_completion:.1f}% of them, and "
        f"delivery split roughly {league_left_pct:.0f}% from the left / {100 - league_left_pct:.0f}% from the right. "
        f"<b>{top_row['team']}</b> were the tournament's most prolific crossing side "
        f"({top_row['crosses_per_match']:.1f} crosses/match). Most crosses league-wide are struck from wide, "
        f"advanced areas close to the by-line rather than the half-spaces or deep positions.</p>"
        f"<p><b>Select exactly one team</b> in the Team slicer above to swap in a tailored coaching summary and "
        f"tactical instruction for that opponent.</p>"
        f"{CAVEAT_HTML}"
    )

    for team in sorted(ts["team"].unique()):
        row = ts[ts["team"] == team].iloc[0]
        sub = all_cr[all_cr["team"] == team]
        crosses = int(row["total_crosses"])
        if crosses == 0:
            summaries[team] = (
                f"<p><b>{team}</b> did not register any recorded crosses in this dataset across "
                f"{int(row['matches_played'])} matches.</p>{CAVEAT_HTML}"
            )
            continue

        left_pct = round(row["left_side"] / crosses * 100, 1)
        right_pct = round(100 - left_pct, 1)
        dominant_side = "left" if left_pct >= right_pct else "right"
        dom_pct = max(left_pct, right_pct)

        combo = sub.groupby(["cross_channel", "cross_depth"]).size().sort_values(ascending=False)
        top_channel, top_depth = combo.index[0]
        top_combo_pct = round(combo.iloc[0] / crosses * 100, 1)

        half_counts = sub["half_label"].value_counts()
        half1, half2 = int(half_counts.get("1st Half", 0)), int(half_counts.get("2nd Half", 0))
        half_note = ""
        if half1 and half2:
            more_half = "1st half" if half1 > half2 else "2nd half"
            share = round(max(half1, half2) / (half1 + half2) * 100)
            half_note = f" Crossing volume leans toward the {more_half} ({share:.0f}% of in-game crosses)."

        players = sub.groupby("player").size().sort_values(ascending=False)
        top_player, top_player_n = players.index[0], int(players.iloc[0])
        top_player_pct = round(top_player_n / crosses * 100, 1)
        if top_player_pct >= 25:
            source_note = (
                f" Delivery is notably concentrated in one player \u2014 <b>{top_player}</b> alone accounts for "
                f"{top_player_pct:.0f}% of the team's crosses, making him the clear individual focus for pressing "
                f"and man-marking out of games."
            )
        else:
            source_note = (
                f" <b>{top_player}</b> is the top source ({top_player_n} crosses, {top_player_pct:.0f}% of the "
                f"team total), but supply is shared across multiple players rather than one outlet."
            )

        summary = (
            f"<p><b>{team}</b> crossed the ball {crosses} times across {int(row['matches_played'])} matches "
            f"({row['crosses_per_match']:.1f} per match), completing {row['completion_pct']:.1f}% of deliveries. "
            f"{dom_pct:.0f}% of crosses come from the <b>{dominant_side} side</b>, and the most common delivery "
            f"pattern is {top_depth.lower()} from the {top_channel.lower()} ({top_combo_pct:.0f}% of crosses)."
            f"{half_note}{source_note} {row['box_entry_pct']:.0f}% of crosses reach the penalty area, and "
            f"{int(row['dangerous'])} were both completed and delivered into the box \u2014 the closest proxy "
            f"this dataset has for a genuine goal-scoring chance.</p>"
            f"<p><b>Tactical instruction for the back line:</b> shift defensive coverage toward the "
            f"{dominant_side} flank, engage <b>{top_player}</b> early to cut off the primary supply line, and "
            f"prioritize denying the {top_channel.lower()} at {top_depth.lower()} range before the ball reaches "
            f"the by-line or the box.</p>"
            f"{CAVEAT_HTML}"
        )
        summaries[team] = summary
    return summaries



def build_cross_records(all_cr):
    """Row-level cross data (~1,200 rows) embedded once; every slicer/chart in the
    browser filters and re-aggregates this client-side, mirroring the Excel slicers."""
    d = all_cr.copy()
    d["c"] = d["completed"] == "Yes"
    d["box"] = d["box_entry"] == "Yes"
    d["dgr"] = d["dangerous_cross"] == "Yes"
    d["x"] = d["x"].round(1)
    d["y"] = d["y"].round(1)
    out = d[[
        "team", "match_label", "half_label", "side", "player", "x", "y",
        "x_band", "y_band", "cross_channel", "cross_depth", "c", "box", "dgr",
    ]].rename(columns={
        "match_label": "match", "half_label": "half", "x_band": "xb", "y_band": "yb",
        "cross_channel": "ch", "cross_depth": "dp",
    })
    return out.to_json(orient="records")


def build_pass_agg(df):
    """Passes aggregated to (team, match, half) so the 'Total Passes' KPI and the
    'Matches' KPI can be filtered client-side without shipping all 53,888 rows."""
    g = df.groupby(["team", "match_label", "half_label"]).agg(
        total=("match_id", "size"), completed=("pass_completed_flag", "sum")
    ).reset_index().rename(columns={"match_label": "match", "half_label": "half"})
    return g.to_json(orient="records")


def build_report():
    print("Loading data...")
    df = load_enriched_passes()
    ts = team_summary(df)
    all_cr = df[df["cross"] == True].copy()
    teams_list = sorted(all_cr["team"].unique())
    matches_list = sorted(df["match_label"].unique())

    print("Building row-level payloads for client-side filtering...")
    cross_records_json = build_cross_records(all_cr)
    pass_agg_json = build_pass_agg(df)
    coach_summaries_json = json.dumps(build_coach_summaries(df, all_cr, ts))

    # ---------------- League-wide comparison charts (static categories) -------
    ts_sorted = ts.sort_values("crosses_per_match", ascending=True)
    leaderboard_order = ts_sorted["team"].tolist()
    fig_leaderboard = go.Figure(go.Bar(
        y=leaderboard_order, x=ts_sorted["crosses_per_match"], orientation="h",
        marker_color=BLUE, name="Crosses / Match",
        text=ts_sorted["crosses_per_match"], texttemplate="%{text:.1f}", textposition="outside",
    ))
    fig_leaderboard.update_layout(title="Crosses per Match — All 24 Teams", height=650,
                                   xaxis_title="Crosses / Match", yaxis_title=None, showlegend=False)

    ts_comp = ts.sort_values("completion_pct", ascending=True)
    completion_order = ts_comp["team"].tolist()
    fig_completion = go.Figure(go.Bar(
        y=completion_order, x=ts_comp["completion_pct"], orientation="h",
        marker_color=GREEN, name="Completion %",
        text=ts_comp["completion_pct"], texttemplate="%{text:.1f}%", textposition="outside",
    ))
    fig_completion.update_layout(title="Cross Completion % — All 24 Teams", height=650,
                                  xaxis_title="Completion %", yaxis_title=None, showlegend=False)

    ts_side = ts.sort_values("total_crosses", ascending=False)
    side_order = ts_side["team"].tolist()
    fig_side = go.Figure()
    fig_side.add_trace(go.Bar(x=side_order, y=ts_side["left_side"], name="Left Side", marker_color=BLUE))
    fig_side.add_trace(go.Bar(x=side_order, y=ts_side["right_side"], name="Right Side", marker_color=GOLD))
    fig_side.update_layout(barmode="stack", title="Left vs Right Crossing Volume by Team", height=460,
                            xaxis_title=None, yaxis_title="Crosses", xaxis_tickangle=-45)

    # ---------------- Zone heatmaps / delivery / half / pitch / players (initial, unfiltered) ----------------
    def heatmap_fig(z, title, div_id):
        fig = go.Figure(data=go.Heatmap(
            z=z, x=[f"{c}-{c+20}" for c in ZONE_X], y=[f"{r}-{r+16}" for r in ZONE_Y],
            colorscale=[[0, "#FFFFFF"], [0.5, "#FFD96B"], [1, "#C0392B"]], showscale=True,
            hovertemplate="X band %{x}<br>Y band %{y}<br>Crosses: %{z}<extra></extra>",
        ))
        fig.update_layout(title=title, height=340, xaxis_title="Pitch length (x) — own half → by-line",
                           yaxis_title="Pitch width (y) — left → right")
        return fig

    completed_cr = all_cr[all_cr["completed"] == "Yes"]
    fig_zone_all = heatmap_fig(zone_matrix(all_cr), "Zone Heatmap — All Crosses", "zone_all")
    fig_zone_comp = heatmap_fig(zone_matrix(completed_cr), "Zone Heatmap — Completed Crosses Only", "zone_comp")

    delivery0 = delivery_matrix(all_cr)
    fig_delivery = go.Figure(data=go.Heatmap(
        z=delivery0, x=DEPTH_ORDER, y=CHANNEL_ORDER,
        colorscale=[[0, "#FFFFFF"], [0.5, "#8FBFEA"], [1, "#0071E3"]], showscale=True,
        text=delivery0, texttemplate="%{text}",
        hovertemplate="%{y} / %{x}<br>Crosses: %{z}<extra></extra>",
    ))
    fig_delivery.update_layout(title="Delivery Profile — Channel × Depth", height=340)

    total_h, completed_h = half_data(all_cr)
    fig_half = go.Figure()
    fig_half.add_trace(go.Bar(x=HALF_ORDER, y=total_h, name="Total Crosses", marker_color=NAVY,
                               text=total_h, texttemplate="%{text}", textposition="outside"))
    fig_half.add_trace(go.Bar(x=HALF_ORDER, y=completed_h, name="Completed", marker_color=GREEN,
                               text=completed_h, texttemplate="%{text}", textposition="outside"))
    fig_half.update_layout(barmode="group", title="Crosses by Match Period", height=400, yaxis_title="Crosses")

    p_labels, p_total, p_completed = players_data(all_cr, 15, show_team_suffix=True)
    fig_players = go.Figure(go.Bar(
        y=p_labels, x=p_total, orientation="h", marker_color=NAVY, name="Total Crosses",
        text=p_completed, texttemplate="%{text} completed", textposition="outside",
    ))
    fig_players.update_layout(title="Top Crossers — Euro 2024", height=560, xaxis_title="Total Crosses",
                               showlegend=False)

    comp0 = all_cr[all_cr["completed"] == "Yes"]
    incomp0 = all_cr[all_cr["completed"] == "No"]
    fig_pitch = go.Figure()
    fig_pitch.add_trace(go.Scatter(
        x=comp0["x"], y=comp0["y"], mode="markers", name=f"Completed ({len(comp0)})",
        text=comp0["player"], hovertemplate="<b>%{text}</b><br>x=%{x}, y=%{y}<br>Completed<extra></extra>",
        marker=dict(color=GREEN, size=8, symbol="circle", line=dict(width=1, color="white")),
    ))
    fig_pitch.add_trace(go.Scatter(
        x=incomp0["x"], y=incomp0["y"], mode="markers", name=f"Incomplete ({len(incomp0)})",
        text=incomp0["player"], hovertemplate="<b>%{text}</b><br>x=%{x}, y=%{y}<br>Incomplete<extra></extra>",
        marker=dict(color=RED, size=8, symbol="x"),
    ))
    fig_pitch.update_layout(
        title=dict(text="Cross Origin Pitch Map — All Teams", x=0, xanchor="left"),
        shapes=pitch_shapes(),
        xaxis=dict(range=[-2, 122], showgrid=False, zeroline=False, title="Pitch length (own half → goal)"),
        yaxis=dict(range=[82, -2], showgrid=False, zeroline=False, title="Pitch width (left → right)"),
        height=520, legend=dict(orientation="h", yanchor="bottom", y=1.06, xanchor="left", x=0),
    )

    # ---------------- League Overview table ----------------
    ts_table = ts.rename(columns={
        "team": "Team", "matches_played": "MP", "total_passes": "Passes", "total_crosses": "Crosses",
        "crosses_per_match": "CPM", "completion_pct": "Completion",
        "left_side": "Left", "right_side": "Right", "box_entry_pct": "BoxPct", "dangerous": "Dangerous",
    })
    # ranked by volume of passes (most passes first) by default; column headers are sortable client-side
    ts_table = ts_table.sort_values("Passes", ascending=False).reset_index(drop=True)
    table_rows = "".join(
        f'<tr data-team="{r.Team}"><td>{r.Team}</td><td>{r.MP}</td><td>{r.Passes:,}</td><td>{r.Crosses}</td>'
        f"<td>{r.CPM:.2f}</td><td>{r.Completion:.1f}%</td><td>{r.Left}</td><td>{r.Right}</td>"
        f"<td>{r.BoxPct:.1f}%</td><td>{r.Dangerous}</td></tr>"
        for r in ts_table.itertuples()
    )
    league_json = ts_table.rename(columns={
        "Team": "team", "MP": "mp", "Passes": "passes", "Crosses": "crosses", "CPM": "cpm",
        "Completion": "completion", "Left": "left", "Right": "right", "BoxPct": "boxpct", "Dangerous": "dangerous",
    }).to_json(orient="records")

    total_matches = int(df["match_id"].nunique())
    total_passes = int(len(df))
    total_crosses = int(len(all_cr))
    completion0 = round((all_cr["completed"] == "Yes").mean() * 100, 1) if total_crosses else 0.0
    box0 = round((all_cr["box_entry"] == "Yes").mean() * 100, 1) if total_crosses else 0.0
    dangerous0 = int((df["dangerous_cross"] == "Yes").sum())

    kpi_html = "".join(
        f'<div class="kpi-card"><div class="kpi-value" id="kpi-{k}">{v}</div><div class="kpi-label">{l}</div>'
        f'<div class="kpi-sub">{s}</div></div>'
        for k, l, v, s in [
            ("matches", "Matches", total_matches, "51 group + knockout fixtures"),
            ("passes", "Total Passes", f"{total_passes:,}", "All 24 teams, every minute"),
            ("crosses", "Total Crosses", f"{total_crosses:,}", "Volume for current selection"),
            ("completion", "Cross Completion", f"{completion0:.1f}%", "Possession retained after delivery"),
            ("box", "Crosses Targeting Box", f"{box0:.1f}%", "End point lands inside the penalty area"),
            ("dangerous", "Dangerous Crosses", dangerous0, "Completed AND reached the box"),
        ]
    )

    def tiles_html(container_id, values):
        return "".join(f'<button class="tile" data-value="{v}">{v}</button>' for v in values)

    slicer_html = dict(
        team_tiles=tiles_html("tiles-team", teams_list),
        match_tiles=tiles_html("tiles-match", matches_list),
        half_tiles=tiles_html("tiles-half", HALF_ORDER),
        side_tiles=tiles_html("tiles-side", ["Left", "Right"]),
    )

    plots_html = dict(
        pitch=fig_to_div(fig_pitch, "pitch"),
        zone_all=fig_to_div(fig_zone_all, "zone_all"),
        zone_comp=fig_to_div(fig_zone_comp, "zone_comp"),
        delivery=fig_to_div(fig_delivery, "delivery"),
        half=fig_to_div(fig_half, "half"),
        leaderboard=fig_to_div(fig_leaderboard, "leaderboard"),
        completion=fig_to_div(fig_completion, "completion"),
        side=fig_to_div(fig_side, "side"),
        players=fig_to_div(fig_players, "players"),
    )

    orders_json = json.dumps(dict(leaderboard=leaderboard_order, completion=completion_order, side=side_order))

    html = HTML_TEMPLATE.format(
        kpi_html=kpi_html, table_rows=table_rows,
        cross_records_json=cross_records_json, pass_agg_json=pass_agg_json, orders_json=orders_json,
        coach_summaries_json=coach_summaries_json, league_json=league_json,
        **slicer_html, **plots_html,
    )
    OUT_FILE.write_text(html, encoding="utf-8")
    print("Saved:", OUT_FILE)



HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Euro 2024 — Cross &amp; Passing Intelligence Report</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
  :root {{
    --navy: #1B2A4A; --ink: #1D1D1F; --muted: #6E6E73; --blue: #0071E3;
    --bg: #F5F5F7; --card: #FFFFFF; --border: #E5E5EA;
  }}
  * {{ box-sizing: border-box; }}
  body {{
    margin: 0; background: var(--bg); color: var(--ink);
    font-family: -apple-system, BlinkMacSystemFont, 'SF Pro Display', 'Segoe UI', Helvetica, Arial, sans-serif;
    -webkit-font-smoothing: antialiased;
  }}
  .hero {{
    background: linear-gradient(180deg, #0B1526 0%, #1B2A4A 60%, #223862 100%);
    color: white; padding: 72px 8vw 56px 8vw; text-align: center;
  }}
  .hero h1 {{ font-size: 44px; font-weight: 700; letter-spacing: -0.02em; margin: 0 0 12px 0; }}
  .hero p {{ font-size: 19px; color: #C9D3E6; max-width: 720px; margin: 0 auto; line-height: 1.5; }}
  .hero .credit {{ margin-top: 28px; font-size: 13px; color: #8CA0C4; letter-spacing: 0.02em; }}
  .filter-bar {{
    position: sticky; top: 0; z-index: 50; background: rgba(255,255,255,0.92); backdrop-filter: blur(14px);
    border-bottom: 1px solid var(--border); padding: 16px 8vw; box-shadow: 0 2px 16px rgba(0,0,0,0.04);
  }}
  .filter-bar .filter-top {{
    display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px;
    margin-bottom: 10px;
  }}
  .filter-bar .now-showing {{ font-size: 13px; color: var(--muted); }}
  .filter-bar .now-showing b {{ color: var(--blue); }}
  .reset-all {{
    font-size: 12px; font-weight: 700; color: white; background: var(--navy); border: none;
    border-radius: 20px; padding: 6px 14px; cursor: pointer;
  }}
  .reset-all:hover {{ background: #10192E; }}
  .slicer-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; }}
  .slicer-panel {{
    background: white; border: 1px solid var(--border); border-radius: 12px; padding: 8px 10px;
  }}
  .slicer-panel .slicer-head {{
    display: flex; align-items: center; justify-content: space-between; margin-bottom: 6px;
  }}
  .slicer-panel .slicer-title {{
    font-size: 10px; font-weight: 700; color: var(--muted); text-transform: uppercase; letter-spacing: 0.05em;
  }}
  .slicer-panel .clear-btn {{
    font-size: 10px; color: var(--blue); background: none; border: none; cursor: pointer; padding: 0;
    font-weight: 600;
  }}
  .slicer-tiles {{
    display: flex; flex-wrap: wrap; gap: 5px; max-height: 92px; overflow-y: auto; padding-right: 2px;
  }}
  .tile {{
    font-size: 11.5px; font-weight: 500; color: var(--ink); background: #F0F0F3; border: 1px solid transparent;
    border-radius: 7px; padding: 5px 9px; cursor: pointer; transition: all 0.12s ease; white-space: nowrap;
  }}
  .tile:hover {{ background: #E4E4EA; }}
  .tile.selected {{ background: var(--blue); color: white; }}
  .tile.unavailable {{ display: none; }}
  .section {{ max-width: 1180px; margin: 0 auto; padding: 64px 24px; }}
  .section h2 {{ font-size: 28px; font-weight: 700; letter-spacing: -0.01em; margin: 0 0 8px 0; color: var(--navy); }}
  .section .lede {{ color: var(--muted); font-size: 15px; max-width: 760px; margin: 0 0 32px 0; line-height: 1.6; }}
  .kpi-grid {{
    display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px;
    max-width: 1180px; margin: 40px auto 0 auto; padding: 0 24px 0 24px; position: relative; z-index: 2;
  }}
  .kpi-card {{
    background: var(--card); border-radius: 18px; padding: 24px 20px; box-shadow: 0 8px 30px rgba(0,0,0,0.08);
    border: 1px solid var(--border); transition: box-shadow 0.2s ease;
  }}
  .kpi-value {{ font-size: 32px; font-weight: 700; color: var(--navy); letter-spacing: -0.02em; }}
  .kpi-label {{ font-size: 13px; font-weight: 600; color: var(--ink); margin-top: 6px; }}
  .kpi-sub {{ font-size: 12px; color: var(--muted); margin-top: 4px; }}
  .card {{
    background: var(--card); border-radius: 18px; padding: 24px; box-shadow: 0 4px 20px rgba(0,0,0,0.05);
    border: 1px solid var(--border); margin-bottom: 24px;
  }}
  .grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 24px; }}
  @media (max-width: 900px) {{ .grid-2 {{ grid-template-columns: 1fr; }} }}
  table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  th, td {{ text-align: right; padding: 10px 12px; border-bottom: 1px solid var(--border); }}
  th:first-child, td:first-child {{ text-align: left; }}
  th {{ color: var(--muted); font-weight: 600; text-transform: uppercase; font-size: 11px; letter-spacing: 0.03em; }}
  #leagueTable th {{ cursor: pointer; user-select: none; white-space: nowrap; }}
  #leagueTable th:hover {{ color: var(--blue); }}
  #leagueTable th .sort-arrow {{ display: inline-block; width: 12px; color: var(--blue); font-size: 9px; }}
  tr:hover {{ background: #FAFAFC; }}
  tr.highlight {{ background: #EEF3FF; }}
  tr.highlight td:first-child {{ font-weight: 700; color: var(--blue); }}
  .footer {{ text-align: center; padding: 48px 24px 80px 24px; color: var(--muted); font-size: 13px; }}
  .footer a {{ color: var(--blue); text-decoration: none; }}
  .tag {{ display:inline-block; background:#EEF3FF; color:var(--blue); font-size:11px; font-weight:700;
          padding:4px 10px; border-radius:20px; letter-spacing:0.03em; margin-bottom:14px; }}
  .coach-card {{
    background: linear-gradient(180deg, #FFFFFF 0%, #F7F9FF 100%); border: 1px solid var(--border);
    border-radius: 20px; padding: 32px 36px; box-shadow: 0 8px 30px rgba(0,0,0,0.06);
  }}
  .coach-card h3 {{
    margin: 0 0 14px 0; font-size: 13px; font-weight: 700; color: var(--blue); text-transform: uppercase;
    letter-spacing: 0.05em;
  }}
  .coach-card p {{ font-size: 15px; line-height: 1.7; color: var(--ink); margin: 0 0 14px 0; }}
  .coach-card p:last-child {{ margin-bottom: 0; }}
  .coach-card p.caveat {{
    font-size: 12.5px; color: var(--muted); background: #FFF7E6; border-left: 3px solid #C9A227;
    padding: 10px 14px; border-radius: 8px; line-height: 1.5;
  }}
</style>
</head>
<body>

<div class="hero">
  <div class="tag">EURO 2024 &middot; STATSBOMB OPEN DATA</div>
  <h1>Cross &amp; Passing Intelligence</h1>
  <p>A coach-facing breakdown of every cross and pass across all 51 Euro 2024 matches &mdash;
     built to answer: how much does each team cross, from where, how often does it work,
     and who is doing it.</p>
  <div class="credit">Data source: StatsBomb open data via github.com/hudl/open-data &middot; free for non-commercial use</div>
</div>

<div class="filter-bar">
  <div class="filter-top">
    <span class="now-showing">Filter with the slicers below (click a tile to toggle it, like Excel slicers) &mdash; every chart, KPI, table row and the coach's summary update instantly. Now showing: <b id="nowShowing">All Teams &middot; All Matches &middot; All Periods &middot; All Sides</b></span>
    <button class="reset-all" id="resetAll">Reset all filters</button>
  </div>
  <div class="slicer-grid">
    <div class="slicer-panel">
      <div class="slicer-head"><span class="slicer-title">Team</span><button class="clear-btn" data-dim="team">Clear</button></div>
      <div class="slicer-tiles" id="tiles-team">{team_tiles}</div>
    </div>
    <div class="slicer-panel">
      <div class="slicer-head"><span class="slicer-title">Match (date order)</span><button class="clear-btn" data-dim="match">Clear</button></div>
      <div class="slicer-tiles" id="tiles-match">{match_tiles}</div>
    </div>
    <div class="slicer-panel">
      <div class="slicer-head"><span class="slicer-title">Period</span><button class="clear-btn" data-dim="half">Clear</button></div>
      <div class="slicer-tiles" id="tiles-half">{half_tiles}</div>
    </div>
    <div class="slicer-panel">
      <div class="slicer-head"><span class="slicer-title">Side</span><button class="clear-btn" data-dim="side">Clear</button></div>
      <div class="slicer-tiles" id="tiles-side">{side_tiles}</div>
    </div>
  </div>
</div>

<div class="kpi-grid">
  {kpi_html}
</div>

<div class="section">
  <h2>Team Deep Dive — Pitch Map</h2>
  <p class="lede">Use the slicers above to see exactly where a team's crosses were struck from.
     Green circles are completed deliveries (possession retained); red crosses are lost.
     Hover any point to see the player who played it.</p>
  <div class="card">{pitch}</div>
</div>

<div class="section">
  <h2>Zone Heatmaps</h2>
  <p class="lede">Cross origin density across a 6&times;5 pitch grid &mdash; goal on the right,
     left side of the attack at the top (same convention as the pitch map above).
     The completed-only grid isolates deliveries that stayed alive.</p>
  <div class="grid-2">
    <div class="card">{zone_all}</div>
    <div class="card">{zone_comp}</div>
  </div>
</div>

<div class="section">
  <h2>Delivery Profile</h2>
  <p class="lede">Where crosses are struck from: touchline vs half-space (channel) crossed with deep vs advanced vs by-line (depth),
     and how delivery volume shifts across match periods.</p>
  <div class="grid-2">
    <div class="card">{delivery}</div>
    <div class="card">{half}</div>
  </div>
</div>

<div class="section">
  <h2>League-Wide Comparison</h2>
  <p class="lede">How does the selected team's crossing output compare to the rest of the tournament field?
     Selected team(s) are highlighted in navy/gold; the rest of the field stays visible in gray for context.</p>
  <div class="grid-2">
    <div class="card">{leaderboard}</div>
    <div class="card">{completion}</div>
  </div>
  <div class="card">{side}</div>
</div>

<div class="section">
  <h2>Top Crossers</h2>
  <p class="lede">Is delivery concentrated in one player the opposition can scout and stop?</p>
  <div class="card">{players}</div>
</div>

<div class="section">
  <h2>Coach's Briefing</h2>
  <p class="lede">Select exactly one team in the Team slicer above to swap in a tailored tactical summary
     for that opponent — how much they cross, from which side, what delivery pattern works, who to mark,
     and what to tell the back line.</p>
  <div class="coach-card">
    <h3 id="coachTitle">Tournament Overview</h3>
    <div id="coachBody"></div>
  </div>
</div>

<div class="section">
  <h2>League Overview — Full Leaderboard</h2>
  <p class="lede">Ranked by total passes by default &mdash; click any column header to sort by that metric instead (click again to reverse the direction).</p>
  <div class="card">
    <table id="leagueTable">
      <thead><tr>
        <th data-col="team" data-type="text">Team<span class="sort-arrow"></span></th>
        <th data-col="mp">MP<span class="sort-arrow"></span></th>
        <th data-col="passes">Passes<span class="sort-arrow"></span></th>
        <th data-col="crosses">Crosses<span class="sort-arrow"></span></th>
        <th data-col="cpm">Crosses/Match<span class="sort-arrow"></span></th>
        <th data-col="completion">Completion %<span class="sort-arrow"></span></th>
        <th data-col="left">Left<span class="sort-arrow"></span></th>
        <th data-col="right">Right<span class="sort-arrow"></span></th>
        <th data-col="boxpct">Box-Targeting %<span class="sort-arrow"></span></th>
        <th data-col="dangerous">Dangerous<span class="sort-arrow"></span></th>
      </tr></thead>
      <tbody>{table_rows}</tbody>
    </table>
  </div>
</div>

<div class="footer">
  Built for match preparation &middot; ELVTR Soccer Performance Analytics Guest Lecture (Jordyn Kaplan) &middot;
  Data: <a href="https://github.com/hudl/open-data">StatsBomb open data</a>
</div>

<script id="cross-records" type="application/json">{cross_records_json}</script>
<script id="pass-agg" type="application/json">{pass_agg_json}</script>
<script id="chart-orders" type="application/json">{orders_json}</script>
<script id="coach-summaries" type="application/json">{coach_summaries_json}</script>
<script id="league-data" type="application/json">{league_json}</script>
<script>
(function() {{
  const crossRecords = JSON.parse(document.getElementById('cross-records').textContent);
  const passAgg = JSON.parse(document.getElementById('pass-agg').textContent);
  const orders = JSON.parse(document.getElementById('chart-orders').textContent);
  const coachSummaries = JSON.parse(document.getElementById('coach-summaries').textContent);
  const leagueData = JSON.parse(document.getElementById('league-data').textContent);

  const HIGHLIGHT = '#1B2A4A', HIGHLIGHT2 = '#C9A227', DIM = '#D2D2D7';
  const BLUE = '#0071E3', GREEN = '#2EA059', GOLD = '#C9A227';
  const HALF_ORDER = ['1st Half', '2nd Half', 'Extra Time 1', 'Extra Time 2'];
  const ZONE_Y = [64, 48, 32, 16, 0];
  const ZONE_X = [0, 20, 40, 60, 80, 100];
  const CHANNEL_ORDER = ['Wide Channel', 'Half-Space', 'Central'];
  const DEPTH_ORDER = ['Deep', 'Advanced', 'By-line'];

  const state = {{ team: new Set(), match: new Set(), half: new Set(), side: new Set() }};
  const leagueSort = {{ col: 'passes', dir: -1 }};

  function inSet(set, val) {{ return set.size === 0 || set.has(val); }}

  function filterCrosses() {{
    return crossRecords.filter(r =>
      inSet(state.team, r.team) && inSet(state.match, r.match) &&
      inSet(state.half, r.half) && inSet(state.side, r.side)
    );
  }}
  function filterPassAgg() {{
    return passAgg.filter(r => inSet(state.team, r.team) && inSet(state.match, r.match) && inSet(state.half, r.half));
  }}

  function zoneMatrix(rows) {{
    const z = ZONE_Y.map(() => ZONE_X.map(() => 0));
    for (const r of rows) {{
      const ri = ZONE_Y.indexOf(r.yb), ci = ZONE_X.indexOf(r.xb);
      if (ri >= 0 && ci >= 0) z[ri][ci]++;
    }}
    return z;
  }}
  function deliveryMatrix(rows) {{
    const z = CHANNEL_ORDER.map(() => DEPTH_ORDER.map(() => 0));
    for (const r of rows) {{
      const ri = CHANNEL_ORDER.indexOf(r.ch), ci = DEPTH_ORDER.indexOf(r.dp);
      if (ri >= 0 && ci >= 0) z[ri][ci]++;
    }}
    return z;
  }}
  function halfBreakdown(rows) {{
    const total = HALF_ORDER.map(() => 0), completed = HALF_ORDER.map(() => 0);
    for (const r of rows) {{
      const i = HALF_ORDER.indexOf(r.half);
      if (i >= 0) {{ total[i]++; if (r.c) completed[i]++; }}
    }}
    return {{ total, completed }};
  }}
  function playersBreakdown(rows, n, showTeamSuffix) {{
    const map = new Map();
    for (const r of rows) {{
      const key = r.player + '|' + r.team;
      if (!map.has(key)) map.set(key, {{ player: r.player, team: r.team, total: 0, completed: 0 }});
      const o = map.get(key); o.total++; if (r.c) o.completed++;
    }}
    let arr = Array.from(map.values()).filter(o => o.total > 0);
    arr.sort((a, b) => b.total - a.total);
    arr = arr.slice(0, n);
    arr.sort((a, b) => a.total - b.total);
    return {{
      labels: arr.map(o => showTeamSuffix ? `${{o.player}} (${{o.team}})` : o.player),
      total: arr.map(o => o.total), completed: arr.map(o => o.completed),
    }};
  }}
  function pitchArrays(rows) {{
    const comp = rows.filter(r => r.c), incomp = rows.filter(r => !r.c);
    return {{
      cx: comp.map(r => r.x), cy: comp.map(r => r.y), cplayer: comp.map(r => r.player),
      ix: incomp.map(r => r.x), iy: incomp.map(r => r.y), iplayer: incomp.map(r => r.player),
      ccount: comp.length, icount: incomp.length,
    }};
  }}
  function computeKpis(crossRows, passRows) {{
    const crosses = crossRows.length;
    const completed = crossRows.filter(r => r.c).length;
    const box = crossRows.filter(r => r.box).length;
    const dangerous = crossRows.filter(r => r.dgr).length;
    const passes = passRows.reduce((s, r) => s + r.total, 0);
    const matches = new Set(passRows.map(r => r.match)).size;
    return {{
      matches, passes, crosses,
      completion: crosses ? Math.round(completed / crosses * 1000) / 10 : 0,
      box: crosses ? Math.round(box / crosses * 1000) / 10 : 0,
      dangerous,
    }};
  }}
  function colorArray(order, base, highlight) {{
    if (state.team.size === 0) return order.map(() => base);
    return order.map(t => state.team.has(t) ? highlight : DIM);
  }}
  function fmtInt(n) {{ return n.toLocaleString(); }}
  function summaryLabel(dim, label) {{
    const set = state[dim];
    if (set.size === 0) return `All ${{label}}`;
    if (set.size === 1) return [...set][0];
    return set.size + ' ' + label + ' selected';
  }}

  function setTileAvailability(containerId, availableSet) {{
    document.getElementById(containerId).querySelectorAll('.tile').forEach(btn => {{
      btn.classList.toggle('unavailable', !availableSet.has(btn.dataset.value));
    }});
  }}

  function refreshSlicerAvailability() {{
    // Team/Match/Period availability comes from ALL passes (passAgg), so a team's
    // full match list shows up even if it had zero crosses in a given match.
    const teamAvail = new Set(passAgg.filter(r => inSet(state.match, r.match) && inSet(state.half, r.half)).map(r => r.team));
    const matchAvail = new Set(passAgg.filter(r => inSet(state.team, r.team) && inSet(state.half, r.half)).map(r => r.match));
    const halfAvail = new Set(passAgg.filter(r => inSet(state.team, r.team) && inSet(state.match, r.match)).map(r => r.half));
    // Side is cross-specific, so its availability comes from the cross records.
    const sideAvail = new Set(crossRecords.filter(r =>
      inSet(state.team, r.team) && inSet(state.match, r.match) && inSet(state.half, r.half)
    ).map(r => r.side));

    setTileAvailability('tiles-team', teamAvail);
    setTileAvailability('tiles-match', matchAvail);
    setTileAvailability('tiles-half', halfAvail);
    setTileAvailability('tiles-side', sideAvail);
  }}

  function update() {{
    refreshSlicerAvailability();
    const crossRows = filterCrosses();
    const passRows = filterPassAgg();
    const kpis = computeKpis(crossRows, passRows);

    document.getElementById('nowShowing').innerHTML =
      `<b>${{summaryLabel('team', 'Teams')}}</b> &middot; ${{summaryLabel('match', 'Matches')}} &middot; ` +
      `${{summaryLabel('half', 'Periods')}} &middot; ${{summaryLabel('side', 'Sides')}}`;

    document.getElementById('kpi-matches').textContent = fmtInt(kpis.matches);
    document.getElementById('kpi-passes').textContent = fmtInt(kpis.passes);
    document.getElementById('kpi-crosses').textContent = fmtInt(kpis.crosses);
    document.getElementById('kpi-completion').textContent = kpis.completion.toFixed(1) + '%';
    document.getElementById('kpi-box').textContent = kpis.box.toFixed(1) + '%';
    document.getElementById('kpi-dangerous').textContent = fmtInt(kpis.dangerous);

    const pitch = pitchArrays(crossRows);
    const pitchTitle = state.team.size === 0 ? 'All Teams'
      : (state.team.size === 1 ? [...state.team][0] : state.team.size + ' Teams Selected');
    Plotly.restyle('pitch', {{
      x: [pitch.cx, pitch.ix], y: [pitch.cy, pitch.iy], text: [pitch.cplayer, pitch.iplayer],
      name: ['Completed (' + pitch.ccount + ')', 'Incomplete (' + pitch.icount + ')'],
    }}, [0, 1]);
    Plotly.relayout('pitch', {{'title.text': 'Cross Origin Pitch Map — ' + pitchTitle}});

    const completedRows = crossRows.filter(r => r.c);
    Plotly.restyle('zone_all', {{ z: [zoneMatrix(crossRows)] }}, [0]);
    Plotly.restyle('zone_comp', {{ z: [zoneMatrix(completedRows)] }}, [0]);
    const dm = deliveryMatrix(crossRows);
    Plotly.restyle('delivery', {{ z: [dm], text: [dm] }}, [0]);

    const hb = halfBreakdown(crossRows);
    Plotly.restyle('half', {{ y: [hb.total], text: [hb.total] }}, [0]);
    Plotly.restyle('half', {{ y: [hb.completed], text: [hb.completed] }}, [1]);

    const oneTeam = state.team.size === 1;
    const pb = playersBreakdown(crossRows, oneTeam ? 10 : 15, !oneTeam);
    Plotly.restyle('players', {{ y: [pb.labels], x: [pb.total], text: [pb.completed] }}, [0]);
    Plotly.relayout('players', {{'title.text': oneTeam ? 'Top 10 Crossers — ' + [...state.team][0] : 'Top Crossers — Euro 2024'}});

    Plotly.restyle('leaderboard', {{'marker.color': [colorArray(orders.leaderboard, BLUE, HIGHLIGHT2)]}}, [0]);
    Plotly.restyle('completion', {{'marker.color': [colorArray(orders.completion, GREEN, HIGHLIGHT2)]}}, [0]);
    Plotly.restyle('side', {{'marker.color': [colorArray(orders.side, BLUE, HIGHLIGHT)]}}, [0]);
    Plotly.restyle('side', {{'marker.color': [colorArray(orders.side, GOLD, HIGHLIGHT2)]}}, [1]);

    applyLeagueHighlight();

    const coachKey = oneTeam ? [...state.team][0] : 'ALL';
    document.getElementById('coachTitle').textContent = oneTeam ? coachKey + ' — Tactical Summary' : 'Tournament Overview';
    document.getElementById('coachBody').innerHTML = coachSummaries[coachKey] || coachSummaries['ALL'];
  }}

  function applyLeagueHighlight() {{
    document.querySelectorAll('#leagueTable tbody tr').forEach(row => {{
      row.classList.toggle('highlight', state.team.size > 0 && state.team.has(row.dataset.team));
    }});
  }}

  function renderLeague() {{
    const rows = [...leagueData].sort((a, b) => {{
      const va = a[leagueSort.col], vb = b[leagueSort.col];
      const cmp = typeof va === 'string' ? va.localeCompare(vb) : va - vb;
      return leagueSort.dir * cmp;
    }});
    document.querySelector('#leagueTable tbody').innerHTML = rows.map(r => (
      '<tr data-team="' + r.team + '">' +
      '<td>' + r.team + '</td><td>' + r.mp + '</td><td>' + r.passes.toLocaleString() + '</td>' +
      '<td>' + r.crosses + '</td><td>' + r.cpm.toFixed(2) + '</td><td>' + r.completion.toFixed(1) + '%</td>' +
      '<td>' + r.left + '</td><td>' + r.right + '</td><td>' + r.boxpct.toFixed(1) + '%</td>' +
      '<td>' + r.dangerous + '</td></tr>'
    )).join('');
    document.querySelectorAll('#leagueTable thead th').forEach(th => {{
      const arrow = th.querySelector('.sort-arrow');
      arrow.textContent = th.dataset.col === leagueSort.col ? (leagueSort.dir === 1 ? '\u25b2' : '\u25bc') : '';
    }});
    applyLeagueHighlight();
  }}

  document.querySelectorAll('#leagueTable thead th').forEach(th => {{
    th.addEventListener('click', () => {{
      const col = th.dataset.col;
      if (leagueSort.col === col) {{ leagueSort.dir *= -1; }}
      else {{ leagueSort.col = col; leagueSort.dir = th.dataset.type === 'text' ? 1 : -1; }}
      renderLeague();
    }});
  }});

  function makeTiles(containerId, dim) {{
    document.getElementById(containerId).querySelectorAll('.tile').forEach(btn => {{
      btn.addEventListener('click', () => {{
        const v = btn.dataset.value;
        if (state[dim].has(v)) {{ state[dim].delete(v); btn.classList.remove('selected'); }}
        else {{ state[dim].add(v); btn.classList.add('selected'); }}
        update();
      }});
    }});
  }}
  makeTiles('tiles-team', 'team');
  makeTiles('tiles-match', 'match');
  makeTiles('tiles-half', 'half');
  makeTiles('tiles-side', 'side');

  document.querySelectorAll('.clear-btn').forEach(btn => {{
    btn.addEventListener('click', () => {{
      const dim = btn.dataset.dim;
      state[dim].clear();
      document.querySelectorAll('#tiles-' + dim + ' .tile').forEach(b => b.classList.remove('selected'));
      update();
    }});
  }});
  document.getElementById('resetAll').addEventListener('click', () => {{
    ['team', 'match', 'half', 'side'].forEach(dim => {{
      state[dim].clear();
      document.querySelectorAll('#tiles-' + dim + ' .tile').forEach(b => b.classList.remove('selected'));
    }});
    update();
  }});

  renderLeague();
  update();
}})();
</script>


</body>
</html>
"""


if __name__ == "__main__":
    build_report()
