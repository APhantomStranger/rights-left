#!/usr/bin/env python3
"""
newsletter.py — optional pipeline stage (NEWSLETTER).

Runs after Ingest + Rebuild, as a step in the Ingest workflow. Takes the
single most recent week's entries — the exact same data the site itself
shows, read via the same read_workbook()/group_by_week() build_site.py
uses — and creates a DRAFT weekly digest email in Buttondown via its API.

It only ever creates a draft. Nothing is ever sent automatically: you
review it in your Buttondown dashboard and press Send yourself when
you're happy with it. (Buttondown's API defaults a new email to draft
status unless you explicitly ask it to send, so there's no extra flag
needed to keep this safe — the safe behavior is just what happens if you
do nothing special.)

If Ingest runs more than once for the same week (approving a few items
now, more later), this creates a separate draft each time rather than
trying to merge them — simplest to reason about, and any extra drafts
are one click to delete in Buttondown before you send the one you
actually want.

Requires the BUTTONDOWN_API_KEY secret. Without it, or if the API call
fails for any reason, this script prints a notice and exits cleanly — it
never fails the workflow, since a newsletter hiccup should never block
the actual site/workbook update from being committed.

Usage:
    python scripts/newsletter.py --xlsx data/Trump_Second_Term_Weekly_Tracker.xlsx
"""

import argparse, os, sys

try:
    import requests
except ImportError:
    sys.exit("requests not installed — run: pip install -r requirements.txt")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_site import read_workbook, group_by_week  # reuse the site's own logic

BUTTONDOWN_API = "https://api.buttondown.com/v1/emails"
SITE_URL = "https://rightsleft.org"


def format_digest(week_label, entries, site_url):
    """Markdown body for the digest — Buttondown auto-detects and renders
    Markdown, so no HTML needed. Leads with srcdesc (always populated by
    Gather) rather than the AI-dependent event/impact fields, same
    reasoning as the site itself; impact is a bonus line only when present.
    """
    n = len(entries)
    lines = [f"# This week on Rights Left — {week_label}", "",
             f"{n} new entr{'y' if n == 1 else 'ies'} added to the tracker this week.", ""]
    for e in entries:
        lines.append(f"**{e['srcdesc']}**  ")
        if e.get("impact"):
            lines.append(e["impact"] + "  ")
        src_bits = [b for b in (e.get("outlet"), e.get("srcdate")) if b]
        src_line = " — ".join(src_bits)
        if e.get("url"):
            lines.append((f"{src_line} · " if src_line else "") + f"[Read the source]({e['url']})")
        elif src_line:
            lines.append(src_line)
        lines.append("")
    lines += ["---", f"[View the full tracker]({site_url})"]
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--site-url", default=SITE_URL)
    args = ap.parse_args()

    key = os.environ.get("BUTTONDOWN_API_KEY")
    if not key:
        print("No BUTTONDOWN_API_KEY set — skipping newsletter draft. "
              "The site and workbook are unaffected either way.")
        return

    entries = read_workbook(args.xlsx)
    if not entries:
        print("Workbook has no entries — nothing to draft.")
        return

    latest = group_by_week(entries)[0]
    subject = f"Rights Left — week of {latest['week']}"
    body = format_digest(latest["week"], latest["entries"], args.site_url)

    try:
        resp = requests.post(
            BUTTONDOWN_API,
            headers={"Authorization": f"Token {key}"},
            json={"subject": subject, "body": body},
            timeout=30,
        )
    except Exception as ex:
        print(f"  ! could not reach Buttondown ({ex.__class__.__name__}) — skipping draft this run.")
        return

    if resp.status_code >= 300:
        print(f"  ! Buttondown API error {resp.status_code}: {resp.text[:300]}")
        print("  Skipping the newsletter draft this run — nothing else is affected.")
        return

    email_id = (resp.json() or {}).get("id", "?")
    print(f'Drafted "{subject}" in Buttondown (id {email_id}). '
          f"Review and send it from your Buttondown dashboard when you're ready.")


if __name__ == "__main__":
    main()
