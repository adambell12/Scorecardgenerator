#!/usr/bin/env python3
"""
Scorecard Generator

Inputs:
  inputs/ytd/YTD Mon YYYY.xlsx   (latest and previous for movement)
  inputs/monthly/Mon YYYY.xlsx   (monthly snapshots for agent month-vs-month table)

Output:
  outputs/Scorecard - Mon YYYY.pdf
"""
from __future__ import annotations

import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

def get_base_dir() -> Path:
    """Return the project directory.

    When packaged with PyInstaller (especially --onefile), __file__ points into a temporary
    _MEI folder. We want the folder where the executable lives so config.json and the
    inputs/outputs folders can be found next to it.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


@dataclass
class MetricTierRule:
    name: str
    max_time_seconds: Optional[int] = None    # for time metrics (lower better)
    min_value: Optional[float] = None         # for percent metrics (higher better)

@dataclass
class MetricSpec:
    key: str
    label: str
    type: str  # "time" or "percent"
    lower_is_better: bool
    tiers: List[MetricTierRule]

def parse_hhmmss_to_seconds(v: Any) -> Optional[int]:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, pd.Timedelta):
        return int(v.total_seconds())
    # Excel may pass as datetime.time, or string like "0:12:09"
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return None
    # Sometimes extracted as "0 days 00:12:09"
    m = re.search(r'(\d+):(\d{2}):(\d{2})$', s)
    if not m:
        # try mm:ss
        m2 = re.search(r'(\d+):(\d{2})$', s)
        if m2:
            mm, ss = int(m2.group(1)), int(m2.group(2))
            return mm*60 + ss
        return None
    hh, mm, ss = int(m.group(1)), int(m.group(2)), int(m.group(3))
    return hh*3600 + mm*60 + ss

def seconds_to_hhmmss(sec: Optional[int]) -> str:
    if sec is None:
        return "—"
    sec = int(sec)
    hh = sec // 3600
    rem = sec % 3600
    mm = rem // 60
    ss = rem % 60
    return f"{hh}:{mm:02d}:{ss:02d}"

def pct_to_str(v: Any) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    return f"{float(v)*100:.2f}%"

def pts_to_str(v: float) -> str:
    sign = "+" if v >= 0 else ""
    return f"{sign}{v:.2f} pts"

def load_config(cfg_path: Path) -> Dict[str, Any]:
    with open(cfg_path, "r") as f:
        return json.load(f)

def build_metric_specs(cfg: Dict[str, Any]) -> List[MetricSpec]:
    specs: List[MetricSpec] = []
    for m in cfg["metrics"]:
        tiers: List[MetricTierRule] = []
        if m["type"] == "time":
            for t in m["tiers"]:
                mx = t.get("max")
                mxs = parse_hhmmss_to_seconds(mx) if mx else None
                tiers.append(MetricTierRule(name=t["name"], max_time_seconds=mxs))
        else:
            for t in m["tiers"]:
                tiers.append(MetricTierRule(name=t["name"], min_value=t.get("min")))
        specs.append(MetricSpec(
            key=m["key"], label=m["label"], type=m["type"],
            lower_is_better=bool(m["lower_is_better"]),
            tiers=tiers
        ))
    return specs

def tier_for_value(spec: MetricSpec, raw_value: Any) -> str:
    if spec.type == "time":
        sec = parse_hhmmss_to_seconds(raw_value)
        if sec is None:
            return "—"
        for t in spec.tiers:
            if t.max_time_seconds is None:
                return t.name
            if sec <= t.max_time_seconds:
                return t.name
        return spec.tiers[-1].name
    else:
        if raw_value is None or (isinstance(raw_value, float) and pd.isna(raw_value)):
            return "—"
        v = float(raw_value)
        for t in spec.tiers:
            if t.min_value is None:
                return t.name
            if v >= t.min_value:
                return t.name
        return spec.tiers[-1].name

def goal_delta(spec: MetricSpec, raw_value: Any) -> str:
    """
    For time metrics: show how much time to reach next better tier.
    For percent metrics: show points to reach next better tier.
    """
    tier = tier_for_value(spec, raw_value)
    if tier == "—":
        return "—"
    if spec.type == "time":
        sec = parse_hhmmss_to_seconds(raw_value)
        # tiers ordered best->worst in config for time (Excelling, Performing, ...)
        names = [t.name for t in spec.tiers]
        idx = names.index(tier)
        if idx == 0:
            return "At/above goal"
        # next better is idx-1
        target = spec.tiers[idx-1].max_time_seconds
        if target is None or sec is None:
            return "—"
        diff = max(sec - target, 0)
        return f"{seconds_to_hhmmss(diff)} to {spec.tiers[idx-1].name}"
    else:
        v = float(raw_value) if raw_value is not None and not (isinstance(raw_value,float) and pd.isna(raw_value)) else None
        if v is None:
            return "—"
        # percent tiers are best->worst (Exceptional, Solid, ...)
        names = [t.name for t in spec.tiers]
        idx = names.index(tier)
        if idx == 0:
            return "At/above goal"
        target = spec.tiers[idx-1].min_value
        if target is None:
            return "—"
        pts = (target - v) * 100
        if pts <= 0.00001:
            return "At/above goal"
        return f"{pts:.2f} pts to {spec.tiers[idx-1].name}"

def parse_month_from_filename(name: str) -> Tuple[int,int]:
    """
    Accepts:
      'Jan 2026.xlsx'
      'YTD Jan 2026.xlsx'
    Returns (year, month_index 1-12)
    """
    m = re.search(r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})', name)
    if not m:
        raise ValueError(f"Unrecognized month in filename: {name}")
    mon = MONTHS.index(m.group(1)) + 1
    yr = int(m.group(2))
    return (yr, mon)

def find_latest_ytd(ytd_dir: Path) -> Tuple[Path, Optional[Path], str]:
    ytd_files = [p for p in ytd_dir.glob("YTD *.xlsx") if p.is_file()]
    if not ytd_files:
        raise FileNotFoundError("No YTD files found in inputs/ytd (expected 'YTD Mon YYYY.xlsx').")
    ytd_files.sort(key=lambda p: parse_month_from_filename(p.name))
    latest = ytd_files[-1]
    prev = ytd_files[-2] if len(ytd_files) >= 2 else None
    yr, mon = parse_month_from_filename(latest.name)
    title = f"{MONTHS[mon-1]} {yr}"
    return latest, prev, title

def load_snapshot(path: Path, supervisor_name: str, campaign_filter: Optional[str]) -> pd.DataFrame:
    df = pd.read_excel(path, engine="openpyxl")
    # filter campaign if requested
    if campaign_filter:
        df = df[df["Campaign"].astype(str).str.strip() == campaign_filter]
    # filter supervisor if present
    if "Sup_Name_All_Agents" in df.columns:
        df = df[df["Sup_Name_All_Agents"].astype(str).str.strip() == supervisor_name]
    # remove totals and blanks
    df = df[df["Agent"].notna()]
    df = df[df["Agent"].astype(str).str.strip().str.lower() != "total"]
    # normalize agent name
    df["Agent"] = df["Agent"].astype(str).str.strip()
    return df.reset_index(drop=True)

def build_score(df: pd.DataFrame, specs: List[MetricSpec]) -> pd.Series:
    """
    Composite score: tier points + small value bonus.
    Higher is better.
    """
    tier_points_time = {"Excelling":4, "Performing":3, "Approaching":2, "Improvement Needed":1, "—":0}
    tier_points_pct  = {"Exceptional":4, "Solid":3, "Approaching":2, "Improvement Needed":1, "—":0}
    scores = []
    for _, row in df.iterrows():
        s = 0.0
        for spec in specs:
            t = tier_for_value(spec, row.get(spec.key))
            if spec.type == "time":
                s += tier_points_time.get(t, 0)
                sec = parse_hhmmss_to_seconds(row.get(spec.key))
                # bonus for faster times
                if sec is not None:
                    s += max(0.0, (15*60 - sec) / (15*60)) * 0.25
            else:
                s += tier_points_pct.get(t, 0)
                v = row.get(spec.key)
                if v is not None and not (isinstance(v,float) and pd.isna(v)):
                    s += float(v) * 0.25
        scores.append(s)
    return pd.Series(scores, index=df.index)

def compute_rank_table(ytd_df: pd.DataFrame, specs: List[MetricSpec]) -> pd.DataFrame:
    df = ytd_df.copy()
    df["__score"] = build_score(df, specs)
    df = df.sort_values("__score", ascending=False).reset_index(drop=True)
    df["Rank"] = df.index + 1
    return df

def movement(curr_rank: pd.DataFrame, prev_rank: Optional[pd.DataFrame]) -> pd.DataFrame:
    df = curr_rank.copy()
    if prev_rank is None:
        df["Move"] = "—"
        df["MoveDir"] = "flat"
        return df
    prev_map = dict(zip(prev_rank["Agent"], prev_rank["Rank"]))
    moves = []
    dirs = []
    for a, r in zip(df["Agent"], df["Rank"]):
        pr = prev_map.get(a)
        if pr is None:
            moves.append("—")
            dirs.append("new")
        else:
            m = pr - r
            if m > 0:
                moves.append(f"+{m}")
                dirs.append("up")
            elif m < 0:
                moves.append(f"{m}")
                dirs.append("down")
            else:
                moves.append("0")
                dirs.append("flat")
    df["Move"] = moves
    df["MoveDir"] = dirs
    return df

def arrow_for_delta(delta: Any) -> str:
    if delta is None:
        return "→"
    if isinstance(delta, (int,float)):
        if abs(delta) < 1e-9:
            return "→"
        return "▲" if delta > 0 else "▼"
    return "→"

def arrow_for_time_delta(sec_delta: Optional[int], lower_is_better: bool) -> str:
    if sec_delta is None or sec_delta == 0:
        return "→"
    # if lower is better, negative delta (faster) is good (▲)
    if lower_is_better:
        return "▲" if sec_delta < 0 else "▼"
    return "▲" if sec_delta > 0 else "▼"

def fmt_delta_time(sec_delta: Optional[int]) -> str:
    if sec_delta is None:
        return "—"
    sign = "+" if sec_delta >= 0 else "-"
    return f"{sign}{seconds_to_hhmmss(abs(sec_delta))}"

def fmt_delta_pts(pts: Optional[float]) -> str:
    if pts is None:
        return "—"
    sign = "+" if pts >= 0 else ""
    return f"{sign}{pts:.2f} pts"

def safe(v: Any) -> Any:
    return None if (isinstance(v,float) and pd.isna(v)) else v

def build_styles():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="H1", parent=styles["Heading1"], fontSize=16, leading=18, spaceAfter=8))
    styles.add(ParagraphStyle(name="H2", parent=styles["Heading2"], fontSize=13, leading=15, spaceAfter=6))
    styles.add(ParagraphStyle(name="Mono", parent=styles["BodyText"], fontName="Courier", fontSize=9, leading=11))
    styles.add(ParagraphStyle(name="Small", parent=styles["BodyText"], fontSize=9, leading=11))
    return styles

def hex_to_color(h: str):
    h = h.lstrip("#")
    r = int(h[0:2],16)/255
    g = int(h[2:4],16)/255
    b = int(h[4:6],16)/255
    return colors.Color(r,g,b)

def tier_bg(tier: str, tier_colors: Dict[str,str]) -> colors.Color:
    h = tier_colors.get(tier)
    if not h:
        return colors.white
    c = hex_to_color(h)
    return colors.Color(c.red, c.green, c.blue, alpha=0.15)

def make_table(data, col_widths, tier_cells=None, tier_colors=None, textcolor_cells=None):
    t = Table(data, colWidths=col_widths, hAlign="LEFT")
    ts = TableStyle([
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),9),
        ("GRID",(0,0),(-1,-1),0.25,colors.lightgrey),
        ("BACKGROUND",(0,0),(-1,0),colors.whitesmoke),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("LEFTPADDING",(0,0),(-1,-1),4),
        ("RIGHTPADDING",(0,0),(-1,-1),4),
        ("TOPPADDING",(0,0),(-1,-1),2),
        ("BOTTOMPADDING",(0,0),(-1,-1),2),
    ])
    if tier_cells and tier_colors:
        for (r,c,tier) in tier_cells:
            ts.add("BACKGROUND",(c,r),(c,r), tier_bg(tier, tier_colors))


    if textcolor_cells:
        for (r,c,clr) in textcolor_cells:
            ts.add("TEXTCOLOR",(c,r),(c,r), clr)

    t.setStyle(ts)
    return t

def main():
    base = get_base_dir()
    cfg = load_config(base/"config.json")
    specs = build_metric_specs(cfg)
    tier_colors = cfg.get("tier_colors", {})

    ytd_latest, ytd_prev, title = find_latest_ytd(base/"inputs"/"ytd")
    # Determine current month from latest file; for the team summary and agent pages we use latest monthly if present.
    supervisor = cfg.get("supervisor_name","")
    campaign = cfg.get("campaign_filter")

    ytd_curr = load_snapshot(ytd_latest, supervisor, campaign)
    ytd_prev_df = load_snapshot(ytd_prev, supervisor, campaign) if ytd_prev else None

    # For month-vs-month, try to load the latest and previous monthly files matching YTD months
    monthly_dir = base/"inputs"/"monthly"
    monthly_files = [p for p in monthly_dir.glob("*.xlsx") if p.is_file() and not p.name.startswith("YTD")]
    monthly_files.sort(key=lambda p: parse_month_from_filename(p.name))
    monthly_latest = monthly_files[-1] if monthly_files else None
    monthly_prev = monthly_files[-2] if len(monthly_files) >= 2 else None

    mon_curr_df = load_snapshot(monthly_latest, supervisor, campaign) if monthly_latest else None
    mon_prev_df = load_snapshot(monthly_prev, supervisor, campaign) if monthly_prev else None

    curr_rank = compute_rank_table(ytd_curr, specs)
    prev_rank = compute_rank_table(ytd_prev_df, specs) if ytd_prev_df is not None else None
    rank_movement = movement(curr_rank, prev_rank)

    out_path = base/"outputs"/f"Scorecard - {title}.pdf"
    styles = build_styles()
    doc = SimpleDocTemplate(str(out_path), pagesize=landscape(letter),
                            leftMargin=cfg["page"]["margin_in"]*inch,
                            rightMargin=cfg["page"]["margin_in"]*inch,
                            topMargin=cfg["page"]["margin_in"]*inch,
                            bottomMargin=cfg["page"]["margin_in"]*inch)

    story = []

    # Page 1: Movement ranking
    if ytd_prev is not None:
        prev_title = f"{MONTHS[parse_month_from_filename(ytd_prev.name)[1]-1]} YTD"
    else:
        prev_title = "Prior YTD"
    curr_title = f"{MONTHS[parse_month_from_filename(ytd_latest.name)[1]-1]} YTD"
    story.append(Paragraph(f"Monthly Movement Ranking - YTD Rank Movement ({prev_title} -> {curr_title})", styles["H1"]))
    headers = ["(#)","Agent","AHT","Hold","Adh","WR","QA","Audit","Move (YTD)"]
    data = [headers]
    tier_cells=[]
    textcolor_cells=[]
    for i,row in rank_movement.iterrows():
        # show only the configured 6 metrics used in the PDF
        vals=[]
        for spec in specs:
            v=row.get(spec.key)
            if spec.type=="time":
                vals.append(seconds_to_hhmmss(parse_hhmmss_to_seconds(v)))
            else:
                vals.append(pct_to_str(v))
        move = row["Move"]
        move_dir = row["MoveDir"]
        arrow = "→"
        if move_dir=="up": arrow="▲"
        elif move_dir=="down": arrow="▼"
        elif move_dir=="new": arrow="•"
        data.append([str(i+1), row["Agent"], *vals, f"{arrow} {move}"])
        r_idx = i+1
        # Movement column text color: green for ▲, red for ▼
        if arrow == "▲":
            textcolor_cells.append((r_idx, 8, colors.green))
        elif arrow == "▼":
            textcolor_cells.append((r_idx, 8, colors.red))
        # tier shading for metrics
        for j,spec in enumerate(specs):
            tier = tier_for_value(spec, row.get(spec.key))
            tier_cells.append((r_idx, 2+j, tier))
    col_widths = [0.4*inch, 2.2*inch] + [0.9*inch]*6 + [0.9*inch]
    story.append(make_table(data, col_widths, tier_cells=tier_cells, tier_colors=tier_colors, textcolor_cells=textcolor_cells))
    story.append(PageBreak())

    # Page 2: Team summary (use current YTD snapshot so numbers reflect current month's YTD)
    summary_df = ytd_curr
    story.append(Paragraph(f"Team Summary - {title}", styles["H1"]))
    headers = ["Agent"] + [spec.label.split(" %")[0] if spec.type=="time" else spec.label for spec in specs]
    data = [headers]
    tier_cells=[]
    for i,row in summary_df.sort_values("Agent").iterrows():
        row_vals=[]
        for spec in specs:
            v=row.get(spec.key)
            if spec.type=="time":
                row_vals.append(seconds_to_hhmmss(parse_hhmmss_to_seconds(v)))
            else:
                row_vals.append(pct_to_str(v))
        data.append([row["Agent"], *row_vals])
        r_idx=len(data)-1
        for j,spec in enumerate(specs):
            tier=tier_for_value(spec, row.get(spec.key))
            tier_cells.append((r_idx, 1+j, tier))
    col_widths=[2.6*inch]+[1.05*inch]*6
    story.append(make_table(data, col_widths, tier_cells=tier_cells, tier_colors=tier_colors))
    story.append(PageBreak())

    # Agent pages
    agents = sorted(set(summary_df["Agent"].tolist()))
    # Maps for quick lookup
    ytd_map = {a: ytd_curr[ytd_curr["Agent"]==a].iloc[0] for a in ytd_curr["Agent"].unique()}
    mon_curr_map = {a: mon_curr_df[mon_curr_df["Agent"]==a].iloc[0] for a in mon_curr_df["Agent"].unique()} if mon_curr_df is not None else {}
    mon_prev_map = {a: mon_prev_df[mon_prev_df["Agent"]==a].iloc[0] for a in mon_prev_df["Agent"].unique()} if mon_prev_df is not None else {}

    # Titles for month compare table
    if monthly_latest is not None:
        curr_month_title = re.search(r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}', monthly_latest.name).group(0)
    else:
        curr_month_title = title
    if monthly_prev is not None:
        prev_month_title = re.search(r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}', monthly_prev.name).group(0)
    else:
        prev_month_title = "Prior Month"

    # YTD span
    ytd_mon = parse_month_from_filename(ytd_latest.name)[1]
    ytd_span = f"Jan-{MONTHS[ytd_mon-1]}"

    for a in agents:
        story.append(Paragraph(a, styles["H1"]))

        # YTD Summary
        story.append(Paragraph(f"Year-To-Date Summary ({ytd_span})", styles["H2"]))
        ytd_row = ytd_map.get(a)
        ytd_data=[["Metric","YTD Value","YTD Tier","Goal Delta"]]
        tier_cells=[]
        if ytd_row is None:
            for spec in specs:
                ytd_data.append([spec.label,"—","—","—"])
        else:
            for spec in specs:
                raw = safe(ytd_row.get(spec.key))
                val = seconds_to_hhmmss(parse_hhmmss_to_seconds(raw)) if spec.type=="time" else pct_to_str(raw)
                tier = tier_for_value(spec, raw)
                gd = goal_delta(spec, raw)
                ytd_data.append([spec.label, val, tier, gd])
                tier_cells.append((len(ytd_data)-1, 2, tier))
        story.append(make_table(ytd_data, [2.1*inch, 1.2*inch, 1.4*inch, 2.4*inch],
                                tier_cells=tier_cells, tier_colors=tier_colors))
        story.append(Spacer(1, 0.15*inch))

        # Month vs Month
        story.append(Paragraph(f"{prev_month_title} vs {curr_month_title} (with Trend & Delta)", styles["H2"]))
        mv_data=[["Metric", f"{prev_month_title} Value", f"{prev_month_title} Tier",
                  f"{curr_month_title} Value", f"{curr_month_title} Tier", "Δ + Arrow"]]
        tier_cells=[]
        textcolor_cells=[]
        prev_row = mon_prev_map.get(a)
        curr_row = mon_curr_map.get(a)
        for spec in specs:
            prev_raw = safe(prev_row.get(spec.key)) if prev_row is not None else None
            curr_raw = safe(curr_row.get(spec.key)) if curr_row is not None else None
            prev_val = seconds_to_hhmmss(parse_hhmmss_to_seconds(prev_raw)) if spec.type=="time" else pct_to_str(prev_raw)
            curr_val = seconds_to_hhmmss(parse_hhmmss_to_seconds(curr_raw)) if spec.type=="time" else pct_to_str(curr_raw)
            prev_tier = tier_for_value(spec, prev_raw) if prev_row is not None else "—"
            curr_tier = tier_for_value(spec, curr_raw) if curr_row is not None else "—"

            if spec.type=="time":
                psec = parse_hhmmss_to_seconds(prev_raw)
                csec = parse_hhmmss_to_seconds(curr_raw)
                dsec = None if (psec is None or csec is None) else (csec - psec)
                arrow = arrow_for_time_delta(dsec, lower_is_better=True)
                delta_str = fmt_delta_time(dsec)
            else:
                pv = None if prev_raw is None or (isinstance(prev_raw,float) and pd.isna(prev_raw)) else float(prev_raw)
                cv = None if curr_raw is None or (isinstance(curr_raw,float) and pd.isna(curr_raw)) else float(curr_raw)
                d = None if (pv is None or cv is None) else (cv - pv) * 100
                # for percent metrics higher is better
                if d is None or abs(d) < 1e-9:
                    arrow = "→"
                else:
                    arrow = "▲" if d > 0 else "▼"
                delta_str = fmt_delta_pts(d)
            mv_data.append([spec.label, prev_val, prev_tier, curr_val, curr_tier, f"{arrow} {delta_str}"])
            r_idx=len(mv_data)-1
            # Delta column text color: green for ▲, red for ▼
            if arrow == "▲":
                textcolor_cells.append((r_idx, 5, colors.green))
            elif arrow == "▼":
                textcolor_cells.append((r_idx, 5, colors.red))
            tier_cells.append((r_idx, 2, prev_tier))
            tier_cells.append((r_idx, 4, curr_tier))

        story.append(make_table(mv_data, [2.1*inch, 1.35*inch, 1.2*inch, 1.35*inch, 1.2*inch, 1.5*inch],
                                tier_cells=tier_cells, tier_colors=tier_colors, textcolor_cells=textcolor_cells))
        story.append(PageBreak())

    doc.build(story)
    print(f"Generated: {out_path}")

if __name__ == "__main__":
    main()
