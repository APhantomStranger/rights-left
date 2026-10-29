#!/usr/bin/env python3
"""
roster.py — refresh the Watchlist in data/Individual_Crimes_Tracker.xlsx.

Pulls two public datasets and merges them into the Watchlist sheet, keeping
every row you or earlier runs added by hand:

  - Every Republican currently in Congress
    (unitedstates/congress-legislators, the open dataset congress.gov data
    projects use)
  - Every sitting federal judge appointed by Donald Trump, plus the
    Republican-appointed Supreme Court justices
    (Federal Judicial Center biographical database)

Optionally merges extra people from a JSON file: --extra file.json, an array
of {"name", "group", "role", "state"?, "aliases"?}.

Run it from the "Refresh watchlist" workflow (Actions tab) after an election
or confirmation wave, or locally:
    python scripts/roster.py
"""

import argparse, csv, io, json, os, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from individuals import TRACKER, load_tracker, save_tracker

LEGISLATORS = "https://unitedstates.github.io/congress-legislators/legislators-current.json"
FJC_JUDGES = "https://www.fjc.gov/sites/default/files/history/judges.csv"


def _get(url):
    import requests
    r = requests.get(url, timeout=60, headers={"User-Agent": "RightsLeft roster refresh"})
    r.raise_for_status()
    return r


def congress_republicans():
    out = []
    for p in _get(LEGISLATORS).json():
        term = p["terms"][-1]
        if term.get("party") != "Republican":
            continue
        n = p["name"]
        name = n.get("official_full") or f"{n.get('first', '')} {n.get('last', '')}".strip()
        aliases = set()
        if n.get("nickname"):
            aliases.add(f"{n['nickname']} {n['last']}")
        if n.get("first") and n.get("last") and f"{n['first']} {n['last']}" != name:
            aliases.add(f"{n['first']} {n['last']}")
        if term["type"] == "sen":
            group, role = "Senate", f"U.S. Senator ({term['state']})"
        else:
            group = "House"
            d = term.get("district")
            role = f"U.S. Rep. ({term['state']}-{'AL' if d in (0, None) else d})"
        out.append({"name": name, "group": group, "role": role, "state": term["state"],
                    "aliases": ", ".join(sorted(aliases)), "origin": "congress-legislators"})
    return out


def judges():
    rows = list(csv.DictReader(io.StringIO(_get(FJC_JUDGES).content.decode("utf-8-sig"))))
    out = []
    for r in rows:
        if r.get("Death Year"):
            continue
        k = max((i for i in range(1, 7) if r.get(f"Court Name ({i})")), default=None)
        if not k or r.get(f"Termination ({k})"):
            continue
        pres = r.get(f"Appointing President ({k})", "")
        court = r.get(f"Court Name ({k})", "")
        is_scotus = "Supreme Court" in court
        if not ("Trump" in pres or (is_scotus and r.get(f"Party of Appointing President ({k})") == "Republican")):
            continue
        name = " ".join(x for x in (r.get("First Name"), r.get("Middle Name"),
                                    r.get("Last Name"), r.get("Suffix")) if x)
        name = " ".join(name.split())
        short = f"{r.get('First Name', '')} {r.get('Last Name', '')}".strip()
        title = "Justice, U.S. Supreme Court" if is_scotus else court
        if r.get(f"Senior Status Date ({k})"):
            title += " (senior status)"
        out.append({"name": short, "group": "Judiciary",
                    "role": f"{title} — appointed by {pres.replace('Donald J. ', '')}",
                    "state": "", "aliases": name if name != short else "", "origin": "FJC"})
    return out


def merge(watch, new):
    by_name = {w["name"].lower(): w for w in watch}
    added = 0
    for n in new:
        cur = by_name.get(n["name"].lower())
        if cur:
            # refresh role/group from the dataset, keep hand-added aliases
            for k in ("group", "role", "state"):
                if n.get(k) and (cur.get("origin") == n.get("origin") or not cur.get(k)):
                    cur[k] = n[k]
            al = {a.strip() for a in (cur.get("aliases") or "").split(",") if a.strip()}
            al |= {a.strip() for a in (n.get("aliases") or "").split(",") if a.strip()}
            cur["aliases"] = ", ".join(sorted(al))
        else:
            watch.append(dict(n))
            by_name[n["name"].lower()] = watch[-1]
            added += 1
    return added


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tracker", default=TRACKER)
    ap.add_argument("--extra", help="JSON file of extra people to add")
    ap.add_argument("--no-congress", action="store_true")
    ap.add_argument("--no-judges", action="store_true")
    args = ap.parse_args()

    matters, watch = load_tracker(args.tracker)
    for w in watch:
        w.pop("_row", None)
    before = len(watch)
    if not args.no_congress:
        c = congress_republicans()
        print(f"Congress: {len(c)} Republicans, {merge(watch, c)} new to the watchlist")
    if not args.no_judges:
        j = judges()
        print(f"Judiciary: {len(j)} judges, {merge(watch, j)} new to the watchlist")
    if args.extra:
        extra = json.load(open(args.extra, encoding="utf-8"))
        for e in extra:
            e.setdefault("origin", "curated")
        print(f"Extra: {len(extra)} people, {merge(watch, extra)} new to the watchlist")
    # Everyone with a matter is on the watchlist, so weekly updates find them
    mm = [{"name": m["name"], "group": m.get("group", ""), "role": m.get("role", ""),
           "state": m.get("state", ""), "origin": "matters"} for m in matters if m.get("name")]
    print(f"From matters: {merge(watch, mm)} new to the watchlist")
    for m in matters:
        m.pop("_row", None)
    save_tracker(matters, watch, args.tracker)
    print(f"Watchlist: {before} -> {len(watch)} people. Saved {args.tracker}")


if __name__ == "__main__":
    main()
