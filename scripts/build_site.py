#!/usr/bin/env python3
"""
build_site.py — Stage 3 of the Rights Left pipeline (REBUILD SITE).

Reads the master workbook and regenerates rights-left.html in the repo root.
Called automatically by the ingest workflow after new rows are appended.

Usage:
    python scripts/build_site.py --xlsx data/Trump_Second_Term_Weekly_Tracker.xlsx
"""

import argparse, json, os, datetime as dt
from openpyxl import load_workbook

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(SCRIPT_DIR, "site_template.html")


def read_workbook(xlsx_path):
    """Extract timeline entries and source data from the workbook."""
    wb = load_workbook(xlsx_path, data_only=True)
    tl = wb["Weekly Timeline"]
    src = wb["Sources"]

    # Build a lookup of source data by ref number
    sources = {}
    for row in range(3, src.max_row + 1):
        ref = src.cell(row=row, column=1).value
        if ref is None:
            continue
        sources[ref] = {
            "outlet": src.cell(row=row, column=2).value or "",
            "srcdesc": src.cell(row=row, column=3).value or "",
            "url": src.cell(row=row, column=4).value or "",
            "srcdate": src.cell(row=row, column=5).value or "",
        }
        # openpyxl may return the hyperlink target instead of cell text for URLs
        hyp = src.cell(row=row, column=4).hyperlink
        if hyp and hyp.target:
            sources[ref]["url"] = hyp.target

    # Read timeline rows, group by week
    entries = []
    for row in range(3, tl.max_row + 1):
        week = tl.cell(row=row, column=1).value
        if week is None or str(week).strip() == "":
            continue
        ref = tl.cell(row=row, column=6).value
        src_data = sources.get(ref, {})
        entries.append({
            "week": str(week).strip(),
            "n": int(ref) if ref else row - 2,
            "date": str(tl.cell(row=row, column=2).value or "").strip(),
            "cat": str(tl.cell(row=row, column=3).value or "").strip(),
            "event": str(tl.cell(row=row, column=4).value or "").strip(),
            "impact": str(tl.cell(row=row, column=5).value or "").strip(),
            "outlet": src_data.get("outlet", ""),
            "srcdesc": src_data.get("srcdesc", ""),
            "url": src_data.get("url", ""),
            "srcdate": src_data.get("srcdate", ""),
            "image": "",
        })

    return entries


def group_by_week(entries):
    """Group entries into week objects, newest first."""
    groups = []
    current = None
    for e in entries:
        entry = {k: v for k, v in e.items() if k != "week"}
        if current is None or current["week"] != e["week"]:
            current = {"week": e["week"], "entries": []}
            groups.append(current)
        current["entries"].append(entry)
    groups.reverse()  # newest first
    return groups


def build_site_json(groups, cats):
    """Build the SITE metadata object."""
    total = sum(len(g["entries"]) for g in groups)
    weeks = len(groups)
    # Date range: oldest week to today
    if groups:
        oldest = groups[-1]["week"]
        newest = groups[0]["week"]
    else:
        oldest = newest = "—"
    today = dt.date.today().strftime("%b %-d, %Y")
    return {
        "name": "Rights Left",
        "tag": "The Rights That Have Left of Us Project",
        "sub": ("Documented actions of the second Trump administration, "
                "logged week by week, with sources. Jan 20, 2025 – present."),
        "range": f"Jan 20, 2025 – {today}",
        "total": total,
        "weeks": weeks,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True, help="Path to the master workbook")
    ap.add_argument("--out", default="rights-left.html",
                    help="Output HTML path (default: rights-left.html in repo root)")
    args = ap.parse_args()

    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(
            f"Template not found at {TEMPLATE_PATH}. "
            "It should be in the same directory as this script.")

    print(f"Reading workbook: {args.xlsx}")
    entries = read_workbook(args.xlsx)
    print(f"  {len(entries)} entries across {len(set(e['week'] for e in entries))} weeks")

    groups = group_by_week(entries)
    cats = sorted({e["cat"] for e in entries if e["cat"]})
    site = build_site_json(groups, cats)

    # Serialize to JSON, escaping </ for safe embedding in <script>
    data_json = json.dumps(groups, ensure_ascii=False).replace("</", "<\\/")
    cats_json = json.dumps(cats, ensure_ascii=False)
    site_json = json.dumps(site, ensure_ascii=False)

    template = open(TEMPLATE_PATH, encoding="utf-8").read()
    html = (template
            .replace("__DATA__", data_json)
            .replace("__CATS__", cats_json)
            .replace("__SITE__", site_json))

    # Sanity check
    for token in ("__DATA__", "__CATS__", "__SITE__"):
        if token in html:
            raise RuntimeError(f"Token {token} was not replaced — template issue")

    with open(args.out, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Wrote {args.out} ({len(html):,} bytes)")
    print(f"  {site['total']} entries, {site['weeks']} weeks, {len(cats)} categories")


if __name__ == "__main__":
    main()
