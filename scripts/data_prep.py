"""
Data preparation for the Euro 2024 Cross Map dashboard.
Loads the raw passes export + teams reference, and derives the analytics
support columns used throughout the Excel dashboard and HTML report.
"""
import pandas as pd
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PASSES_FILE = ROOT / "euro_2024_passes.xlsx"
TEAMS_FILE = ROOT / "teams.xlsx"

HALF_MAP = {1: "1st Half", 2: "2nd Half", 3: "Extra Time 1", 4: "Extra Time 2"}
DEPTH_MAP = {1: "Deep", 2: "Deep", 3: "Deep", 4: "Deep", 5: "Advanced", 6: "By-line"}
CHANNEL_MAP = {1: "Wide Channel", 2: "Half-Space", 3: "Central", 4: "Half-Space", 5: "Wide Channel"}


def load_teams() -> pd.DataFrame:
    teams = pd.read_excel(TEAMS_FILE)
    teams.columns = [c.strip() for c in teams.columns]
    teams = teams.rename(columns={"team_name": "team", "matches played": "matches_played"})
    return teams


def load_enriched_passes() -> pd.DataFrame:
    df = pd.read_excel(PASSES_FILE)
    teams = load_teams()

    # --- support columns (per instructions doc + coach-facing analytics) ---
    df["match_label"] = df["match_date"].dt.strftime("%Y-%m-%d") + "  " + df["home_team"] + " vs " + df["away_team"]
    df["half_label"] = df["period"].map(HALF_MAP)
    df["pitch_third"] = pd.cut(
        df["x"], bins=[-0.1, 40, 80, 120.1],
        labels=["Defensive Third", "Middle Third", "Attacking Third"]
    ).astype(str)

    x_band_idx = (df["x"] // 20 + 1).clip(upper=6).astype(int)
    y_band_idx = (df["y"] // 16 + 1).clip(upper=5).astype(int)
    df["x_band"] = (x_band_idx - 1) * 20          # 0,20,40,60,80,100 -> pivot column header
    df["y_band"] = (y_band_idx - 1) * 16          # 0,16,32,48,64     -> pivot row header
    df["cross_depth"] = x_band_idx.map(DEPTH_MAP)
    df["cross_channel"] = y_band_idx.map(CHANNEL_MAP)

    df["box_entry"] = ((df["end_x"] >= 102) & (df["end_y"] >= 18) & (df["end_y"] <= 62)).map({True: "Yes", False: "No"})
    df["dangerous_cross"] = (
        (df["cross"] == True) & (df["completed"] == "Yes") & (df["box_entry"] == "Yes")
    ).map({True: "Yes", False: "No"})

    # numeric 0/1 flags so PivotTable measures can SUM() conditional counts under shared slicers
    df["cross_flag"] = df["cross"].astype(int)
    df["cross_completed_flag"] = ((df["cross"] == True) & (df["completed"] == "Yes")).astype(int)
    df["cross_box_entry_flag"] = ((df["cross"] == True) & (df["box_entry"] == "Yes")).astype(int)
    df["cross_dangerous_flag"] = (df["dangerous_cross"] == "Yes").astype(int)
    df["pass_completed_flag"] = (df["completed"] == "Yes").astype(int)

    # scatter-plot helper columns (only meaningful for crosses; NaN elsewhere so charts skip them)
    is_cross = df["cross"] == True
    # flipped (80 - y) so the scatter chart's Y axis reads top=left / bottom=right,
    # matching the zone-grid pivot convention (ascending y_band lists Left at the top row)
    df["cross_y_completed"] = (80 - df["y"]).where(is_cross & (df["completed"] == "Yes"))
    df["cross_y_incomplete"] = (80 - df["y"]).where(is_cross & (df["completed"] == "No"))
    df["cross_x_plot"] = df["x"].where(is_cross)

    df = df.merge(teams, on="team", how="left")
    return df


def team_summary(df: pd.DataFrame) -> pd.DataFrame:
    teams = load_teams()
    crosses = df[df["cross"] == True]

    agg = crosses.groupby("team").agg(
        total_crosses=("cross", "size"),
        completed_crosses=("completed", lambda s: (s == "Yes").sum()),
        left_side=("side", lambda s: (s == "Left").sum()),
        right_side=("side", lambda s: (s == "Right").sum()),
        box_entries=("box_entry", lambda s: (s == "Yes").sum()),
        dangerous=("dangerous_cross", lambda s: (s == "Yes").sum()),
    ).reset_index()

    total_passes = df.groupby("team").size().rename("total_passes").reset_index()

    out = teams.merge(agg, on="team", how="left").merge(total_passes, on="team", how="left")
    out = out.fillna(0)
    for c in ["total_crosses", "completed_crosses", "left_side", "right_side", "box_entries", "dangerous", "total_passes"]:
        out[c] = out[c].astype(int)

    out["crosses_per_match"] = (out["total_crosses"] / out["matches_played"]).round(2)
    out["completion_pct"] = (out["completed_crosses"] / out["total_crosses"].replace(0, pd.NA) * 100).round(1).fillna(0)
    out["box_entry_pct"] = (out["box_entries"] / out["total_crosses"].replace(0, pd.NA) * 100).round(1).fillna(0)
    out = out.sort_values("crosses_per_match", ascending=False).reset_index(drop=True)
    return out


def top_players(df: pd.DataFrame, min_crosses: int = 3) -> pd.DataFrame:
    crosses = df[df["cross"] == True]
    agg = crosses.groupby(["team", "player"]).agg(
        total_crosses=("cross", "size"),
        completed_crosses=("completed", lambda s: (s == "Yes").sum()),
    ).reset_index()
    agg = agg[agg["total_crosses"] >= min_crosses]
    agg["completion_pct"] = (agg["completed_crosses"] / agg["total_crosses"] * 100).round(1)
    agg = agg.sort_values("total_crosses", ascending=False).reset_index(drop=True)
    return agg


if __name__ == "__main__":
    df = load_enriched_passes()
    print(df.shape)
    print(df.dtypes)
    ts = team_summary(df)
    print(ts.head(10).to_string())
    tp = top_players(df)
    print(tp.head(10).to_string())
