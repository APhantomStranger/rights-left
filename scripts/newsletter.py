#!/usr/bin/env python3
"""
newsletter.py — optional pipeline stage (NEWSLETTER).

Runs after Ingest + Rebuild, as a step in the Ingest workflow. Takes the
single most recent week's entries — the exact same data the site itself
shows, read via the same read_workbook()/group_by_week() build_site.py
uses — and SENDS the weekly digest email to every subscriber via the
Buttondown API, styled to match the site's own dark/red/blue identity.
No draft, no review step, no manual click: this goes out live the
moment the workflow reaches this step.

That's a deliberate choice, not an oversight — this project explicitly
considered a review-first draft mode and a scheduled-with-a-window mode
before landing here. If that ever needs to change back, the fix is
small: drop the "status": "about_to_send" field (and the two headers
below) from the request, and it reverts to Buttondown's own default of
creating a plain draft instead.

The HTML template below is hand-built for email specifically, NOT a
copy of the site's actual CSS — most email clients (Outlook especially)
don't support CSS variables, web fonts, or modern layout, so this uses
inline styles, web-safe font stacks, and a simple table structure
instead. It won't be pixel-identical to the website; it's designed to
evoke the same identity within what email actually allows.

If Ingest runs more than once for the same week (approving a few items
now, more later), this sends a separate email each time rather than
merging them — there is no undo once a send goes out, so avoid running
Ingest twice against the same week's file once this is live.

Requires the BUTTONDOWN_API_KEY secret. Without it, or if the API call
fails for any reason, this script prints a notice and exits cleanly — it
never fails the workflow, since a newsletter hiccup should never block
the actual site/workbook update from being committed.

Usage:
    python scripts/newsletter.py --xlsx data/Trump_Second_Term_Weekly_Tracker.xlsx
"""

import argparse, html, os, sys

try:
    import requests
except ImportError:
    sys.exit("requests not installed — run: pip install -r requirements.txt")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_site import read_workbook, group_by_week  # reuse the site's own logic

BUTTONDOWN_API = "https://api.buttondown.com/v1/emails"
SITE_URL = "https://rightsleft.org"

# Email-safe palette/fonts — hardcoded hex and web-safe font stacks, not the
# site's CSS variables or Google Fonts, since most email clients support
# neither. See the module docstring for why.
INK, LINE = "#0A0A0B", "#2A2D33"
PAPER, PAPER_DIM, PAPER_FAINT = "#F6F6F4", "#B6B7BD", "#77787F"
RED, BLUE, BLUE_BRIGHT = "#E53935", "#3B6FF6", "#6B93FF"
SANS = "Helvetica,Arial,sans-serif"
MONO = "'Courier New',Courier,monospace"
DISPLAY = "'Arial Black',Arial,Helvetica,sans-serif"


def _esc(s):
    return html.escape(s or "")


def _entry_row(e):
    """One entry as a table row. Every optional field degrades gracefully:
    a blank category hides that line entirely rather than showing an empty
    label, and a missing URL just omits the source link, rather than either
    crashing or leaving a broken/empty link behind.
    """
    cat_html = ""
    if e.get("cat"):
        cat_html = (f'<div style="font-family:{MONO};font-size:10px;letter-spacing:1.5px;'
                    f'color:{BLUE_BRIGHT};text-transform:uppercase;margin:0 0 8px 0;">'
                    f'{_esc(e["cat"])}</div>')
    impact_html = ""
    if e.get("impact"):
        impact_html = (f'<div style="font-family:{SANS};font-size:13px;color:{PAPER_DIM};'
                        f'line-height:1.5;margin:0 0 8px 0;">{_esc(e["impact"])}</div>')
    src_bits = [b for b in (e.get("outlet"), e.get("srcdate")) if b]
    src_line = " &middot; ".join(_esc(b) for b in src_bits)
    link = ""
    if e.get("url"):
        link = (f'<a href="{_esc(e["url"])}" style="color:{BLUE_BRIGHT};text-decoration:none;">'
                f'Read the source &rarr;</a>')
    sep = " &middot; " if src_line and link else ""

    return f'''
    <tr><td style="padding:18px 28px;border-bottom:1px solid {LINE};">
      {cat_html}
      <div style="font-family:{SANS};font-size:16px;font-weight:bold;color:{PAPER};line-height:1.4;margin:0 0 6px 0;">{_esc(e["srcdesc"])}</div>
      {impact_html}
      <div style="font-family:{SANS};font-size:12.5px;color:{PAPER_FAINT};">{src_line}{sep}{link}</div>
    </td></tr>'''


def format_digest(week_label, entries, site_url):
    """Full HTML email body, styled to evoke the site's identity within
    what email clients actually support (see module docstring). The
    leading HTML comment forces Buttondown to treat this as rich HTML
    rather than trying to run it through Markdown conversion.
    """
    n = len(entries)
    rows = "".join(_entry_row(e) for e in entries)
    return f'''<!-- buttondown-editor-mode: fancy -->
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:{INK};padding:36px 0;">
<tr><td align="center">
<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background-color:{INK};">

  <tr><td style="padding:0 28px 26px 28px;border-bottom:2px solid {LINE};">
    <div style="font-family:{MONO};font-size:11px;letter-spacing:3px;color:{PAPER_FAINT};text-transform:uppercase;margin:0 0 14px 0;">Weekly Dispatch</div>
    <div style="font-family:{DISPLAY};font-size:32px;font-weight:900;line-height:1;">
      <span style="color:{RED};">RIGHTS</span><span style="color:{PAPER};">&nbsp;</span><span style="color:{BLUE};">LEFT</span>
    </div>
    <div style="font-family:{SANS};font-size:13px;color:{PAPER_FAINT};margin:10px 0 0 0;">The Rights That Have Left of Us Project</div>
  </td></tr>

  <tr><td style="padding:24px 28px 6px 28px;">
    <div style="font-family:{SANS};font-size:19px;font-weight:bold;color:{PAPER};">Week of {_esc(week_label)}</div>
    <div style="font-family:{SANS};font-size:13.5px;color:{PAPER_DIM};margin:6px 0 0 0;">{n} new entr{'y' if n == 1 else 'ies'} added to the tracker this week.</div>
  </td></tr>

  {rows}

  <tr><td style="padding:26px 28px 6px 28px;">
    <a href="{site_url}" style="font-family:{SANS};font-size:14px;font-weight:bold;color:{BLUE_BRIGHT};text-decoration:none;">View the full tracker &rarr;</a>
  </td></tr>
  <tr><td style="padding:6px 28px 28px 28px;">
    <a href="{site_url}" style="font-family:{SANS};font-size:12.5px;color:{PAPER_FAINT};text-decoration:underline;">See the complete record, every entry ever logged</a>
  </td></tr>

</table>
</td></tr>
</table>'''


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--site-url", default=SITE_URL)
    args = ap.parse_args()

    key = os.environ.get("BUTTONDOWN_API_KEY")
    if not key:
        print("No BUTTONDOWN_API_KEY set — skipping newsletter send. "
              "The site and workbook are unaffected either way.")
        return

    entries = read_workbook(args.xlsx)
    if not entries:
        print("Workbook has no entries — nothing to send.")
        return

    latest = group_by_week(entries)[0]
    subject = f"Rights Left — week of {latest['week']}"
    body = format_digest(latest["week"], latest["entries"], args.site_url)

    try:
        resp = requests.post(
            BUTTONDOWN_API,
            headers={
                "Authorization": f"Token {key}",
                "X-API-Version": "2026-04-01",
                # Confirms this is really meant to send immediately, not just
                # draft. Buttondown only requires this once per API key, but
                # it's harmless to always include, so always including it
                # means this never depends on remembering whether that
                # one-time confirmation already happened.
                "X-Buttondown-Live-Dangerously": "true",
            },
            json={"subject": subject, "body": body, "status": "about_to_send"},
            timeout=30,
        )
    except Exception as ex:
        print(f"  ! could not reach Buttondown ({ex.__class__.__name__}) — skipping send this run.")
        return

    if resp.status_code >= 300:
        print(f"  ! Buttondown API error {resp.status_code}: {resp.text[:300]}")
        print("  Skipping the newsletter send this run — nothing else is affected.")
        return

    email_id = (resp.json() or {}).get("id", "?")
    print(f'Sent "{subject}" via Buttondown (id {email_id}).')


if __name__ == "__main__":
    main()
