#!/usr/bin/env python3
"""
enrich.py — optional Stage 1.5 of the Rights Left pipeline (ENRICH).

Runs after you mark rows include=y in the candidates CSV, but BEFORE
ingest.py appends them into the workbook. For every approved row that's
still missing category / event / impact, this script:

  1. Fetches the actual article at that row's URL
  2. Extracts the article's main text (skipping nav bars, ads, boilerplate)
  3. Asks Claude to pick one of the 17 categories and write a short
     event + impact line, grounded in what the article actually says
  4. Saves the result back into the same CSV, in place

Rows that already have category/event/impact filled in are left completely
untouched — your own wording always wins over the model's. Rows whose
article can't be fetched (paywall, bot-blocking, dead link, no URL) are
left blank with a warning printed to the log; you fill those in by hand
afterward, exactly as ingest.py already allows.

Requires the ANTHROPIC_API_KEY secret. Without it, this script prints a
notice and exits cleanly — the rest of the pipeline still works, you just
fill fields in by hand.

Usage:
    python scripts/enrich.py --csv candidates/2026-07-13.csv
    python scripts/enrich.py                 # uses newest CSV in candidates/
"""

import argparse, csv, json, os, re, sys, glob, time

try:
    import requests
except ImportError:
    sys.exit("requests not installed — run: pip install -r requirements.txt")

try:
    import trafilatura
except ImportError:
    trafilatura = None  # falls back to a crude tag-strip if unavailable

CATEGORIES = [
    "Civil Liberties", "Civil Rights & Minorities", "Courts & SCOTUS",
    "Democracy & Rule of Law", "Economy & Tariffs", "Education",
    "Environment & Science", "Executive Power", "Federal Workforce",
    "Foreign Policy & Aid", "Healthcare", "Immigration",
    "Immigration / Free Speech", "LGBTQ+ Rights", "Press Freedom",
    "Public Health", "Women's Rights / LGBTQ+",
]

FIELDS = ("category", "event", "impact")
TRUTHY = {"y", "yes", "true", "1", "x"}
MAX_ARTICLE_CHARS = 6000
REQUEST_TIMEOUT = 20
HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36")
}


def needs_enrichment(row):
    if (row.get("include") or "").strip().lower() not in TRUTHY:
        return False
    return any(not (row.get(f) or "").strip() for f in FIELDS)


def fetch_article_text(url):
    """Download the page and extract the main article text. Returns None on
    any failure — callers treat that as 'leave this row blank'."""
    if not url:
        return None
    try:
        resp = requests.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
    except Exception as ex:
        print(f"     ! could not fetch article ({ex.__class__.__name__}): {url}")
        return None

    text = None
    if trafilatura is not None:
        try:
            text = trafilatura.extract(resp.text, include_comments=False,
                                        include_tables=False)
        except Exception:
            text = None
    if not text:
        # crude fallback: strip tags, collapse whitespace
        text = re.sub(r"<script.*?</script>|<style.*?</style>", " ",
                       resp.text, flags=re.S | re.I)
        text = re.sub(r"<[^>]+>", " ", text)
        text = re.sub(r"\s+", " ", text).strip()
    if not text or len(text) < 200:
        print(f"     ! article text too short/empty after extraction: {url}")
        return None
    return text[:MAX_ARTICLE_CHARS]


def extract_json(text):
    text = text.strip()
    text = re.sub(r"^```(?:json)?", "", text)
    text = re.sub(r"```$", "", text)
    m = re.search(r"\{.*\}", text.strip(), re.S)
    if m:
        text = m.group(0)
    return json.loads(text)


def classify(client, outlet, headline, article_text):
    cats = ", ".join(CATEGORIES)
    prompt = (
        "You classify US news for a civil-liberties tracker documenting "
        "actions of the Trump administration and their documented effects "
        "on rights, institutions, and the rule of law. Read the article "
        "text below and respond ONLY with JSON in this exact shape, no "
        "preamble, no markdown fences: "
        '{"category": "...", "event": "...", "impact": "..."}. '
        f"category must be exactly one of: {cats}. "
        "event = <=25 words, a plain factual description of what the "
        "administration did — no editorializing. "
        "impact = <=40 words, why courts, experts, or affected communities "
        "describe it as harmful, grounded only in what the article actually "
        "says. If the article is NOT primarily about a specific action taken "
        "by the Trump administration, set category to \"SKIP\" and leave "
        "event and impact as empty strings.\n\n"
        f"Outlet: {outlet}\nHeadline: {headline}\n\n"
        f"Article text:\n{article_text}"
    )
    msg = client.messages.create(
        model="claude-sonnet-4-6", max_tokens=500, temperature=0.2,
        messages=[{"role": "user", "content": prompt}])
    raw = "".join(b.text for b in msg.content if b.type == "text")
    return extract_json(raw)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--csv", help="Candidates CSV to enrich (default: newest in candidates/)")
    ap.add_argument("--candir", default="candidates")
    args = ap.parse_args()

    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        print("No ANTHROPIC_API_KEY set — skipping enrichment. "
              "Approved rows keep whatever category/event/impact you already "
              "typed (or leave blank to fill in by hand later).")
        return

    try:
        import anthropic
    except ImportError:
        print("anthropic package not installed — skipping enrichment.")
        return

    csv_path = args.csv
    if not csv_path:
        pool = sorted(glob.glob(os.path.join(args.candir, "*.csv")))
        if not pool:
            print("No candidates CSV found — nothing to enrich.")
            return
        csv_path = pool[-1]
    print(f"Enriching: {csv_path}")

    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames
        rows = list(reader)

    targets = [r for r in rows if needs_enrichment(r)]
    if not targets:
        print("No approved rows need enrichment (either none are include=y, "
              "or all already have category/event/impact filled in).")
        return
    print(f"{len(targets)} approved row(s) missing details — reading their articles...")

    client = anthropic.Anthropic(api_key=key)
    enriched = 0
    for row in targets:
        url = (row.get("url") or "").strip()
        headline = row.get("srcdesc") or ""
        outlet = row.get("outlet") or ""
        print(f"  -> {headline[:70] or '(no headline)'}")

        article_text = fetch_article_text(url)
        if not article_text:
            print("     leaving blank — fill in by hand later")
            continue

        try:
            data = classify(client, outlet, headline, article_text)
        except Exception as ex:
            print(f"     ! classification failed ({ex.__class__.__name__}) — leaving blank")
            continue

        cat = (data.get("category") or "").strip()
        if not cat or cat == "SKIP":
            print("     model judged this isn't a specific administration action — left blank")
            continue
        if cat not in CATEGORIES:
            print(f"     ! model returned an unrecognized category '{cat}' — left blank")
            continue

        row["category"] = cat
        row["event"] = (data.get("event") or "").strip()
        row["impact"] = (data.get("impact") or "").strip()
        enriched += 1
        print(f"     filled in: {cat}")
        time.sleep(0.4)  # gentle pacing between API calls

    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)

    left_blank = len(targets) - enriched
    print(f"\nEnriched {enriched} row(s)."
          + (f" {left_blank} row(s) still need details filled in by hand."
             if left_blank else "")
          + f" Saved back to {csv_path}.")


if __name__ == "__main__":
    main()
