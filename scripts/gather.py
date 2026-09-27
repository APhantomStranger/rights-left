#!/usr/bin/env python3
"""
gather.py — Stage 1 of the Rights Left pipeline (GATHER).

Two modes:

  gather.py --accumulate --pool-dir pool
      Run every few hours by collect.yml. Pulls ~25 news feeds and saves any
      relevant stories into a rolling per-week pool (the gather-pool branch).
      Needed because most feeds only hold the last day or two of stories.

  gather.py --pool-dir pool
      Run on Sundays by gather.yml. Combines a live pull with the week's pool,
      adds late stories from the end of last week that missed last Sunday's
      sheet, drops anything already offered or already in the tracker, and
      writes a review workbook to  candidates/<monday>.xlsx  — a formatted
      Excel sheet with a dropdown on the Include? column. If that file already
      exists, its rows (and your edits) are kept and only new stories added.

You can add stories yourself in the empty rows under the gathered ones; see the
"How to add your own" tab in the sheet.

You then EDIT that workbook (the "approve" step) in Excel: pick "y" from the
Include? dropdown on every row you want kept (rows marked y highlight green as
you go). Category / event / impact are optional — leave them for enrich.py, or
fill them in yourself if you'd rather. Re-upload the edited file to candidates/
(replacing the original — same filename), then Stage 2 (ingest.py) appends the
approved rows into the tracker workbook automatically.

Optional: set an ANTHROPIC_API_KEY secret and pass --draft to have the model
pre-fill category / event / impact for you (you still review before ingest).
Without --draft, no API key is needed and nothing is sent anywhere.
"""

import os, sys, re, argparse, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

try:
    import feedparser
except ImportError:
    sys.exit("feedparser not installed — run: pip install -r requirements.txt")

# ---- Feeds (edit freely; a feed that errors is skipped, not fatal) -----------
FEEDS = [
    # (outlet name shown on the site, feed URL). Several outlets have both a
    # politics feed and a national/US feed: a lot of what this tracker covers
    # (ICE and police shootings, raids, court fights) runs as national news,
    # not politics, so politics-only feeds were missing it.
    # CNN is gone: its public RSS stopped updating in 2023.
    ("NPR",                 "https://feeds.npr.org/1014/rss.xml"),
    ("NPR",                 "https://feeds.npr.org/1003/rss.xml"),
    ("BBC News",            "https://feeds.bbci.co.uk/news/world/us_and_canada/rss.xml"),
    ("The New York Times",  "https://rss.nytimes.com/services/xml/rss/nyt/Politics.xml"),
    ("The New York Times",  "https://rss.nytimes.com/services/xml/rss/nyt/US.xml"),
    ("The Washington Post", "https://feeds.washingtonpost.com/rss/politics"),
    ("The Washington Post", "https://feeds.washingtonpost.com/rss/national"),
    ("Politico",            "https://rss.politico.com/politics-news.xml"),
    ("NBC News",            "https://feeds.nbcnews.com/nbcnews/public/politics"),
    ("NBC News",            "https://feeds.nbcnews.com/nbcnews/public/news"),
    ("The Guardian",        "https://www.theguardian.com/us-news/us-politics/rss"),
    ("The Guardian",        "https://www.theguardian.com/us-news/rss"),
    ("CBS News",            "https://www.cbsnews.com/latest/rss/politics"),
    ("CBS News",            "https://www.cbsnews.com/latest/rss/us"),
    ("ABC News",            "https://abcnews.go.com/abcnews/politicsheadlines"),
    ("ABC News",            "https://abcnews.go.com/abcnews/usheadlines"),
    ("PBS NewsHour",        "https://www.pbs.org/newshour/feeds/rss/politics"),
    ("PBS NewsHour",        "https://www.pbs.org/newshour/feeds/rss/nation"),
    ("Los Angeles Times",   "https://www.latimes.com/politics/rss2.0.xml"),
    ("Bloomberg",           "https://feeds.bloomberg.com/politics/news.rss"),
    ("Axios",               "https://api.axios.com/feed/"),
    ("The Hill",            "https://thehill.com/homenews/feed/"),
    ("ProPublica",          "https://www.propublica.org/feeds/propublica/main"),
    ("The Intercept",       "https://theintercept.com/feed/?rss"),
    ("Al Jazeera",          "https://www.aljazeera.com/xml/rss/all.xml"),
]

# ---- Relevance filter: keep items whose title/summary mentions any of these --
KEYWORDS = [
    # --- Core administration & politics ---
    "trump", "white house", "executive order", "administration",
    "pardon", "vance", "rfk", "kennedy", "bondi", "patel", "epstein",
    "hegseth", "noem", "shutdown", "voting", "election", "federal",
    "insurrection", "national guard",

    # --- Immigration, ICE & DHS ---
    # (ICE itself is matched separately, as a capitalised whole word — see
    # ICE_RE below. A lowercase "ice " also matched police, justice, office.)
    "i.c.e.", "immigration", "deport", "deportation",
    "customs enforcement", "immigration enforcement",
    "ice agent", "ice raid", "ice arrest", "ice detention",
    "homeland security", "dhs", "border", "border patrol",
    "asylum", "refugee", "migrant", "visa", "birthright",
    "tps", "temporary protected status", "sanctuary",

    # --- Courts & law ---
    "supreme court", "scotus", "doj", "justice dept",
    "justice department",

    # --- Economy ---
    "tariff", "medicaid", "snap",

    # --- Science & public health ---
    "cdc", "nih", "fda", "vaccine", "public health",
    "climate", "environmental protection", "epa",
    "science", "scientific", "research funding",
    "anti-science", "rfk health",

    # --- Education ---
    "education department", "dept of education", "title ix",
    "student loan", "school", "university", "college",
    "academic freedom", "book ban", "curriculum",

    # --- Press, censorship & media manipulation ---
    "fcc", "federal communications commission", "press freedom",
    "journalist", "censor", "media manipulation", "propaganda",
    "disinformation", "state media", "press access",
    "broadcast license", "broadcasting license", "fairness doctrine",
    "license revoked", "pulled off air", "pulled off the air",
    "taken off the air", "show cancelled", "show canceled",
    "network suspends", "talk show host", "late night host",
    "kimmel", "colbert",

    # --- Renaming & historical revisionism ---
    "renamed", "rename", "erase history", "erasing history",
    "rewrite history", "rewriting history", "historical marker",
    "national archives", "smithsonian",

    # --- Law enforcement, force & policing ---
    "police", "law enforcement", "federal agent", "officer",
    "shooting", "fatally shot", "shot and killed", "shot and wounded",
    "shot dead", "use of force", "excessive force", "tear gas",
    "pepper spray", "body camera", "bodycam", "fbi", "u.s. marshals",
    "us marshals", "atf", "dea", "arrest", "detain", "custody", "raid",
    "protester", "crackdown", "consent decree", "civil rights investigation",
]

# Whole-word, case-sensitive: catches "ICE agents shot..." without matching
# "police", "justice", "office", "iced", or a winter-storm "ice".
ICE_RE = re.compile(r"\bICE\b")

# Keywords only match at the START of a word (so "deport" still catches
# "deported"), never in the middle of one ("vance" no longer hits "advance").
_KW_RES = [(k, re.compile(r"(?<![a-z0-9])" + re.escape(k))) for k in KEYWORDS]

# Terms that point straight at the administration or its enforcement arms.
# They weigh more when ranking, so "Trump's arch moves ahead" outranks a local
# story that merely mentions a school and a police officer.
CORE_KEYWORDS = {
    "trump", "white house", "executive order", "administration", "pardon",
    "vance", "rfk", "kennedy", "bondi", "patel", "epstein", "hegseth", "noem",
    "insurrection", "national guard", "i.c.e.", "deport", "deportation",
    "customs enforcement", "immigration enforcement", "ice agent", "ice raid",
    "ice arrest", "ice detention", "homeland security", "dhs", "border patrol",
    "asylum", "birthright", "temporary protected status", "supreme court",
    "scotus", "doj", "justice department", "tariff", "fcc",
    "federal communications commission", "censor", "press freedom",
    "smithsonian", "national archives", "federal agent", "use of force",
    "excessive force", "fatally shot", "shot and killed", "shot and wounded",
    "shot dead",
}

# ---- Semantic concepts (for the OPTIONAL embeddings-based matcher) ------------
# These are rich, meaning-carrying descriptions of what the tracker covers —
# NOT keywords. When an embeddings API key is configured, each candidate
# headline is compared against these by *meaning*, so an article that's about
# one of these themes but shares no literal vocabulary with the KEYWORDS list
# still gets caught. Phrased as full descriptions on purpose: the richer and
# more specific the language here, the better the semantic match.
#
# This never removes anything — an article is kept if it matches a keyword OR
# is semantically close to one of these. It only ever widens the net.
CONCEPTS = [
    "The Trump administration takes an official action expanding executive power or bypassing Congress, the courts, or established legal limits.",
    "Federal immigration enforcement, ICE, or border agents detain, deport, raid, or use force against immigrants, or the administration restricts asylum, visas, refugees, or birthright citizenship.",
    "The administration pressures, threatens, or retaliates against news organizations, broadcasters, or journalists, or a television network pulls, cancels, or suspends a show or host over political content.",
    "Government censorship, propaganda, suppression of speech, or manipulation of media and information by the state.",
    "The FCC or another regulator threatens broadcast licenses, penalizes stations, or uses its authority to influence what networks air.",
    "The administration renames a place, building, landmark, or institution, or alters, removes, or rewrites historical records, monuments, exhibits, or educational content about history.",
    "An action affecting elections, voting access, ballot rules, election administration, or the integrity of the democratic process.",
    "The administration undermines the rule of law, defies a court order, politicizes the Justice Department, or erodes democratic norms and institutions.",
    "A rollback of civil rights or protections for minorities, women, or LGBTQ+ people, or discriminatory federal policy.",
    "Cuts, firings, or political interference in the federal workforce, government agencies, or independent watchdogs and inspectors general.",
    "Attacks on scientific research, public health agencies, environmental protections, vaccines, or the suppression of scientific findings.",
    "Federal action affecting schools, universities, academic freedom, curriculum, book bans, or student funding.",
    "Actions on tariffs, the economy, healthcare, Medicaid, or federal benefits programs that affect ordinary people.",
    "Police, federal agents, or other law enforcement shoot, injure, arrest, detain, or raid people, or face allegations of misconduct, excessive force, or civil-rights violations.",
]

# Similarity threshold: a candidate is kept if its cosine similarity to ANY
# concept above is >= this. Tuned conservatively — high enough to avoid
# sweeping in unrelated political news, low enough to catch genuine reworded
# matches. Raise it if too much junk gets in; lower it if real stories slip by.
SEMANTIC_THRESHOLD = 0.62

# ---- The categories the workbook uses (for the AI drafter / your reference)
CATEGORIES = [
    "Civil Rights & Minorities", "Courts & SCOTUS",
    "Democracy & Rule of Law", "Economy & Tariffs", "Education", "Elections",
    "Environment & Science", "Executive Power", "Federal Workforce",
    "Foreign Policy & Aid", "Free Speech", "Healthcare", "Immigration",
    "Infrastructure and History", "Law Enforcement", "LGBTQ+ Rights",
    "National Security",
    "Press Freedom", "Public Health",
]

# Most NEW candidates one weekly sheet will hold. When more than this match,
# the strongest matches are kept (see score_item), not whichever sorted first.
MAX_CANDIDATES = 250

# Blank rows under the gathered ones that still get the dropdowns, for
# stories you add by hand.
MANUAL_ROWS = 50

# ---- Review workbook layout ---------------------------------------------------
# (row-dict key, column header, column width). Include + Headline lead since
# those are the two things you actually look at to decide keep-or-skip; the
# AI-optional fields are grouped together after; week_of/dates trail since
# they're administrative and rarely need a glance.
REVIEW_COLS = [
    ("include",  "Include?",                  10),
    ("srcdesc",  "Headline / Description",    55),
    ("outlet",   "Outlet",                    16),
    ("srcdate",  "Date",                      12),
    ("url",      "Link",                      40),
    ("category", "Category (optional)",       26),
    ("event",    "Event (optional)",          40),
    ("impact",   "Impact (optional)",         40),
    ("week_of",  "Week Of",                   14),
    ("dates",    "Date(s)",                   12),
]

NAVY, LIGHT_GREEN = "1F3864", "C6E8C6"
HDR_FONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", start_color=NAVY)
BODY = Font(name="Arial", size=10)
LINK_FONT = Font(name="Arial", size=10, color="0563C1", underline="single")
WRAP = Alignment(wrap_text=True, vertical="top")
THIN = Border(bottom=Side(style="thin", color="BFBFBF"))


def write_candidates_xlsx(rows, path, more=()):
    """Write the week's candidates as a formatted, easy-to-review workbook:
    a dropdown on Include?, rows that highlight green once marked y, wrapped
    text so headlines are readable, a clickable Link column, and a frozen,
    filterable header row."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Candidates"

    # Hidden helper sheet holding the 17 categories, purely as a source for
    # the Category dropdown below. Excel's inline data-validation list has a
    # ~255 character limit and the category names together run past that,
    # so the dropdown has to point at a cell range instead — this sheet is
    # that range. You'll never need to open it; it's hidden by default.
    lists_ws = wb.create_sheet("Lists")
    for i, cat in enumerate(CATEGORIES, start=1):
        lists_ws.cell(row=i, column=1, value=cat)
    lists_ws.sheet_state = "hidden"

    n = len(REVIEW_COLS)
    for c, (key, label, width) in enumerate(REVIEW_COLS, start=1):
        cell = ws.cell(row=1, column=c, value=label)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        ws.column_dimensions[get_column_letter(c)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(n)}1"

    url_col = [k for k, _, _ in REVIEW_COLS].index("url") + 1
    for r, row in enumerate(rows, start=2):
        for c, (key, label, width) in enumerate(REVIEW_COLS, start=1):
            cell = ws.cell(row=r, column=c, value=row.get(key) or None)
            cell.font = BODY
            cell.alignment = WRAP
            cell.border = THIN
        if row.get("url"):
            link = ws.cell(row=r, column=url_col)
            link.hyperlink = row["url"]
            link.font = LINK_FONT

    # Dropdowns and the green highlight run well past the last gathered row,
    # so any story you add yourself underneath gets the same Include? and
    # Category dropdowns as the gathered ones.
    last_row = max(len(rows) + 1, 2) + MANUAL_ROWS

    dv = DataValidation(type="list", formula1='"y,n"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"A2:A{last_row}")

    cat_col = get_column_letter([k for k, _, _ in REVIEW_COLS].index("category") + 1)
    dv_cat = DataValidation(
        type="list", formula1=f"Lists!$A$1:$A${len(CATEGORIES)}", allow_blank=True)
    ws.add_data_validation(dv_cat)
    dv_cat.add(f"{cat_col}2:{cat_col}{last_row}")

    green = PatternFill("solid", start_color=LIGHT_GREEN)
    ws.conditional_formatting.add(
        f"A2:{get_column_letter(n)}{last_row}",
        FormulaRule(formula=['LOWER($A2)="y"'], fill=green))

    # Weaker matches that didn't fit under MAX_CANDIDATES. Nothing is thrown
    # away silently: to use one, copy its row onto the Candidates tab and
    # mark it y. Ingest never reads this tab.
    if more:
        mws = wb.create_sheet(MORE_TAB)
        mcols = [c for c in REVIEW_COLS if c[0] in
                 ("srcdesc", "outlet", "srcdate", "url", "week_of", "dates")]
        for c, (key, label, width) in enumerate(mcols, start=1):
            cell = mws.cell(row=1, column=c, value=label)
            cell.font = HDR_FONT; cell.fill = HDR_FILL
            mws.column_dimensions[get_column_letter(c)].width = width
        mws.freeze_panes = "A2"
        for r, row in enumerate(more, start=2):
            for c, (key, _, _) in enumerate(mcols, start=1):
                cell = mws.cell(row=r, column=c, value=row.get(key) or None)
                cell.font = BODY; cell.alignment = WRAP; cell.border = THIN
            if row.get("url"):
                link = mws.cell(row=r, column=[k for k, _, _ in mcols].index("url") + 1)
                link.hyperlink = row["url"]; link.font = LINK_FONT

    # A short how-to on a second tab. The Candidates tab stays first and
    # active, which is the one ingest.py reads.
    how = wb.create_sheet("How to add your own")
    how.column_dimensions["A"].width = 110
    lines = [
        ("Adding a story the gather step missed", True),
        ("", False),
        ("1. On the Candidates tab, go to the first empty row under the gathered stories.", False),
        ("2. Pick y in Include?.", False),
        ("3. Paste the headline into Headline / Description and the article URL into Link.", False),
        ("4. Type the article's date in Date (e.g. Sep 20, 2026 — any normal date format works).", False),
        ("5. Pick a Category if you like. Everything else is optional:", False),
        ("     · Outlet — filled in from the link if left blank (bbc.com → BBC News, npr.org → NPR, …)", False),
        ("     · Week Of and Date(s) — worked out from the Date if left blank", False),
        ("     · Event / Impact — left for the enrich step, or type your own", False),
        ("", False),
        ("Ingest treats your rows exactly like gathered ones. A story dated in an earlier", False),
        ("week lands in that earlier week on the site.", False),
    ]
    for i, (text, bold) in enumerate(lines, start=1):
        c = how.cell(row=i, column=1, value=text)
        c.font = Font(name="Arial", size=11 if bold else 10, bold=bold)
    wb.active = 0

    wb.save(path)


def target_week(today=None):
    """Return (monday_date, start, end_exclusive) for the Mon–Sun week that
    contains YESTERDAY. On a Sunday run that is the week ending that Sunday."""
    today = today or dt.datetime.now(dt.timezone.utc).date()
    ref = today - dt.timedelta(days=1)          # yesterday
    monday = ref - dt.timedelta(days=ref.weekday())
    return monday, monday, monday + dt.timedelta(days=7)


def entry_date(e):
    for key in ("published_parsed", "updated_parsed"):
        t = getattr(e, key, None) or (e.get(key) if hasattr(e, "get") else None)
        if t:
            return dt.date(t.tm_year, t.tm_mon, t.tm_mday)
    return None


def score_item(title, summary):
    """How strongly an item matches the keyword list. 0 = no match.
    A hit in the headline counts double a hit in the summary; ICE in the
    headline counts extra, since that's the core of what this tracks."""
    t, sm = title.lower(), summary.lower()
    score = 0
    for k, rx in _KW_RES:
        w = 2 if k in CORE_KEYWORDS else 1
        if rx.search(t):
            score += 2 * w
        elif rx.search(sm):
            score += w
    if ICE_RE.search(title):
        score += 6
    elif ICE_RE.search(summary):
        score += 2
    return score


def keyword_match(title, summary):
    """The literal-keyword check — fast, free, always runs."""
    return score_item(title, summary) > 0


# --- Optional semantic layer (Voyage embeddings) ------------------------------
# All of this is a no-op unless VOYAGE_API_KEY is set. When it IS set, we embed
# the concept descriptions once, embed each candidate headline, and keep any
# candidate whose meaning is close to a concept even if it shares no keywords.
_concept_vectors = None          # cached concept embeddings (computed once)
_semantic_enabled = None         # tri-state: None=untried, True/False=known

VOYAGE_URL = "https://api.voyageai.com/v1/embeddings"
VOYAGE_MODEL = "voyage-4"


def _embed(texts, key, input_type):
    """Call Voyage's embeddings API for a list of texts. Returns a list of
    vectors, or None on any failure (network, auth, quota) so callers can
    fall back to keyword-only rather than crash the whole gather run."""
    import json as _json
    from urllib import request as _request, error as _error
    payload = _json.dumps({"input": texts, "model": VOYAGE_MODEL,
                           "input_type": input_type}).encode()
    req = _request.Request(VOYAGE_URL, data=payload, headers={
        "Authorization": f"Bearer {key}", "Content-Type": "application/json"})
    try:
        with _request.urlopen(req, timeout=30) as resp:
            data = _json.loads(resp.read())
        return [item["embedding"] for item in data["data"]]
    except (_error.URLError, KeyError, ValueError, TimeoutError) as ex:
        print(f"  ! embeddings call failed ({ex.__class__.__name__}) — "
              f"falling back to keyword-only for this batch.", file=sys.stderr)
        return None


def _cosine(a, b):
    dot = sum(x * y for x, y in zip(a, b))
    na = sum(x * x for x in a) ** 0.5
    nb = sum(y * y for y in b) ** 0.5
    return dot / (na * nb) if na and nb else 0.0


def _ensure_concepts(key):
    """Embed the CONCEPTS list once and cache it. Sets _semantic_enabled based
    on whether it worked, so we only ever try the concept embedding once."""
    global _concept_vectors, _semantic_enabled
    if _semantic_enabled is not None:
        return _semantic_enabled
    vectors = _embed(CONCEPTS, key, input_type="document")
    if vectors is None:
        _semantic_enabled = False
    else:
        _concept_vectors = vectors
        _semantic_enabled = True
    return _semantic_enabled


def semantic_match(title, summary, key):
    """True if the headline is semantically close to any tracked concept.
    Returns False (never raises) if embeddings are unavailable for any reason."""
    if not _ensure_concepts(key):
        return False
    vecs = _embed([f"{title}. {summary}".strip()], key, input_type="query")
    if not vecs:
        return False
    hv = vecs[0]
    return any(_cosine(hv, cv) >= SEMANTIC_THRESHOLD for cv in _concept_vectors)


def relevant(title, summary, voyage_key=None):
    """Keep an article if it matches a keyword OR (when embeddings are
    configured) is semantically close to a tracked concept. Keyword-only
    behavior is exactly preserved when no key is set."""
    if keyword_match(title, summary):
        return True
    if voyage_key:
        return semantic_match(title, summary, voyage_key)
    return False


def week_monday(d):
    return d - dt.timedelta(days=d.weekday())


def week_label(d):
    """'Sep 14, 2026' — the Monday of d's week, in the label format the
    workbook and site use everywhere."""
    return week_monday(d).strftime("%b %-d, %Y")


def fetch_items(start, end, voyage_key=None, skip_urls=()):
    """Pull every feed and return matching items dated in [start, end).

    Each item is the candidate-row dict plus a 'score' used only to decide
    what to keep if a week overflows MAX_CANDIDATES. URLs in skip_urls are
    not re-checked at all (the rolling pool passes the ones it has already
    judged, so semantic matching isn't paid for twice)."""
    seen_url, seen_title, items = set(skip_urls), set(), []
    rejected = []
    semantic_extra = 0
    for outlet, url in FEEDS:
        try:
            feed = feedparser.parse(url)
        except Exception as ex:
            print(f"  ! skipped {outlet} ({url}): {ex}", file=sys.stderr)
            continue
        if not feed.entries:
            print(f"  ! {outlet} returned no items ({url})", file=sys.stderr)
        for e in feed.entries:
            d = entry_date(e)
            if not d or not (start <= d < end):
                continue
            title = re.sub(r"\s+", " ", (e.get("title") or "")).strip()
            summary = re.sub(r"<[^>]+>", " ", e.get("summary", "") or "")
            link = (e.get("link") or "").strip()
            norm = title.lower()
            if not title or link in seen_url or norm in seen_title:
                continue
            score = score_item(title, summary)
            if not score:
                # only reaches the paid API for items the keywords DIDN'T catch
                if voyage_key and semantic_match(title, summary, voyage_key):
                    score = 1
                    semantic_extra += 1
                else:
                    rejected.append(link)
                    seen_url.add(link)
                    continue
            seen_url.add(link); seen_title.add(norm)
            items.append({
                "include": "", "week_of": week_label(d),
                "dates": d.strftime("%b %-d"), "category": "",
                "event": "", "impact": "", "outlet": outlet,
                "srcdesc": title, "url": link,
                "srcdate": d.strftime("%b %-d, %Y"),
                "score": score, "_date": d.isoformat(),
            })
    if voyage_key:
        print(f"  (semantic matching caught {semantic_extra} extra item(s) "
              f"the keywords missed)")
    return items, rejected


# ---- Rolling pool -------------------------------------------------------------
# Most feeds only hold the last day or two of stories (NPR ~27h, PBS ~32h,
# NYT ~37h, WaPo ~42h, measured Sep 2026). A single Sunday pull therefore
# never saw most of the week. collect.yml runs `gather.py --accumulate` every
# few hours and saves what it finds into one JSON file per week on the
# `gather-pool` branch; Sunday's run reads that pool back in.

def _pool_path(pool_dir, monday):
    return os.path.join(pool_dir, f"{monday:%Y-%m-%d}.json")


def load_pool(pool_dir, monday):
    import json
    path = _pool_path(pool_dir, monday)
    if not pool_dir or not os.path.exists(path):
        return {"items": {}, "seen": []}
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        data.setdefault("items", {}); data.setdefault("seen", [])
        return data
    except Exception as ex:
        print(f"  ! couldn't read pool {path}: {ex}", file=sys.stderr)
        return {"items": {}, "seen": []}


def save_pool(pool_dir, monday, data):
    import json
    os.makedirs(pool_dir, exist_ok=True)
    with open(_pool_path(pool_dir, monday), "w", encoding="utf-8") as f:
        json.dump(data, f, indent=1, ensure_ascii=False, sort_keys=True)


def accumulate(pool_dir, voyage_key=None, today=None, keep_weeks=6):
    """Fetch now and add anything new to this week's and last week's pools."""
    today = today or dt.datetime.now(dt.timezone.utc).date()
    this_mon = week_monday(today)
    weeks = [this_mon - dt.timedelta(days=7), this_mon]
    pools = {w: load_pool(pool_dir, w) for w in weeks}
    known = set()
    for data in pools.values():
        known |= set(data["items"]) | set(data["seen"])
    items, rejected = fetch_items(weeks[0], this_mon + dt.timedelta(days=7),
                                  voyage_key, skip_urls=known)
    added = 0
    for it in items:
        w = week_monday(dt.date.fromisoformat(it["_date"]))
        if w in pools and it["url"] not in pools[w]["items"]:
            pools[w]["items"][it["url"]] = it
            pools[w]["seen"].append(it["url"])
            added += 1
    # remember rejects too, so they aren't re-judged every few hours
    for url in rejected:
        pools[this_mon]["seen"].append(url)
    for w, data in pools.items():
        data["seen"] = sorted(set(data["seen"]))
        save_pool(pool_dir, w, data)
    # tidy: drop pool files older than keep_weeks
    cutoff = this_mon - dt.timedelta(weeks=keep_weeks)
    for name in os.listdir(pool_dir):
        try:
            if dt.date.fromisoformat(name[:10]) < cutoff:
                os.remove(os.path.join(pool_dir, name))
        except ValueError:
            pass
    total = sum(len(d["items"]) for d in pools.values())
    print(f"Pool: +{added} new item(s); {total} held across "
          f"{', '.join(f'{w:%b %-d}' for w in weeks)} weeks")
    return added


# ---- What has already been offered for review ---------------------------------

MORE_TAB = "More matches"


def read_sheet_rows(path, tab="Candidates"):
    """Rows of an existing candidates .xlsx/.csv as field dicts (the same
    header mapping ingest.py uses), so a re-run can keep your edits."""
    import csv
    headers = {label.strip().lower(): key for key, label, _ in REVIEW_COLS}
    if path.lower().endswith(".csv"):
        if tab != "Candidates":
            return []
        with open(path, newline="", encoding="utf-8") as f:
            return [dict(r) for r in csv.DictReader(f)]
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    if tab in wb.sheetnames:
        ws = wb[tab]
    elif tab == "Candidates":
        ws = wb.active
    else:
        return []
    cols = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if v and str(v).strip().lower() in headers:
            cols[headers[str(v).strip().lower()]] = c
    rows = []
    for r in range(2, ws.max_row + 1):
        row = {}
        for key, c in cols.items():
            v = ws.cell(row=r, column=c).value
            if isinstance(v, (dt.datetime, dt.date)):
                v = v.strftime("%b %-d, %Y")
            row[key] = "" if v is None else str(v).strip()
        if any(row.values()):
            rows.append(row)
    return rows


def delivered_urls(candir, tracker_xlsx):
    """Every URL already put in front of you: any candidates sheet (pending or
    processed) plus everything already in the tracker workbook. These are
    never offered again, so pulling in last week's late stories can't
    resurface ones you've already seen and passed on."""
    import glob
    urls = set()
    for path in (glob.glob(os.path.join(candir, "*.xlsx")) +
                 glob.glob(os.path.join(candir, "*.csv")) +
                 glob.glob(os.path.join(candir, "processed", "*.xlsx")) +
                 glob.glob(os.path.join(candir, "processed", "*.csv"))):
        try:
            urls |= {r.get("url") for r in read_sheet_rows(path) if r.get("url")}
            urls |= {r.get("url") for r in read_sheet_rows(path, MORE_TAB) if r.get("url")}
        except Exception as ex:
            print(f"  ! couldn't read {path}: {ex}", file=sys.stderr)
    if tracker_xlsx and os.path.exists(tracker_xlsx):
        from openpyxl import load_workbook
        wb = load_workbook(tracker_xlsx, read_only=True, data_only=True)
        if "Sources" in wb.sheetnames:
            for row in wb["Sources"].iter_rows(min_row=3, values_only=True):
                if len(row) >= 4 and row[3]:
                    urls.add(str(row[3]).strip())
    return urls


def collect(monday, start, end, voyage_key=None, pool_dir=None,
            candir="candidates", tracker_xlsx=None):
    """This week's new candidates: a live pull plus the rolling pool, for this
    week AND last week (stories that broke after last Sunday's sheet was
    made would otherwise fall through the gap). Anything already offered in
    an earlier sheet or already in the tracker is left out. Returns rows
    strongest match first; if more than MAX_CANDIDATES match, keeps the
    strongest."""
    prev_monday = monday - dt.timedelta(days=7)
    live, _ = fetch_items(prev_monday, end, voyage_key)
    merged = {it["url"]: it for it in live}
    if pool_dir:
        pooled = 0
        for w in (prev_monday, monday):
            for url, it in load_pool(pool_dir, w)["items"].items():
                if url not in merged:
                    merged[url] = it
                    pooled += 1
        print(f"  (+{pooled} item(s) from the rolling pool that the live "
              f"pull no longer shows)")
    done = delivered_urls(candir, tracker_xlsx)
    fresh = [it for it in merged.values() if it["url"] not in done]
    # dedupe identical headlines across feeds / pool
    by_title = {}
    for it in fresh:
        k = it["srcdesc"].lower()
        if k not in by_title or it["score"] > by_title[k]["score"]:
            by_title[k] = it
    fresh = list(by_title.values())
    late = sum(1 for it in fresh if it["_date"] < monday.isoformat())
    cut = []
    if len(fresh) > MAX_CANDIDATES:
        print(f"  {len(fresh)} matched; keeping the {MAX_CANDIDATES} strongest "
              f"(the rest go on the sheet's More matches tab)")
        fresh.sort(key=lambda it: (-it["score"], it["_date"]))
        fresh, cut = fresh[:MAX_CANDIDATES], fresh[MAX_CANDIDATES:]
    # Strongest matches first, so the top of the sheet is where the likely
    # approvals are and the tail can be skimmed. The header row has a filter
    # arrow on every column if you'd rather sort by Date.
    fresh.sort(key=lambda it: (-it["score"], it["_date"], it["srcdesc"]))
    if late:
        print(f"  ({late} late item(s) from the week of "
              f"{prev_monday:%b %-d} that weren't in last week's sheet)")
    clean = lambda its: [{k: v for k, v in it.items()
                          if not k.startswith("_") and k != "score"} for it in its]
    return clean(fresh), clean(cut)


def ai_draft(rows):
    """Optional: pre-fill category/event/impact. Needs ANTHROPIC_API_KEY."""
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        print("  (--draft set but no ANTHROPIC_API_KEY; leaving fields blank)")
        return rows
    try:
        import anthropic, json
    except ImportError:
        print("  (anthropic package not installed; leaving fields blank)")
        return rows
    client = anthropic.Anthropic(api_key=key)
    cats = ", ".join(CATEGORIES)
    for r in rows:
        prompt = (
            "You classify US news for a civil-liberties tracker of second-term "
            "Trump administration actions. Given one headline, respond ONLY with "
            "JSON: {\"category\":..., \"event\":..., \"impact\":...}. "
            f"category must be exactly one of: {cats}. "
            "event = <=25 words, plainly what the administration did. "
            "impact = <=40 words, why critics/courts/data call it harmful. "
            "If the headline is NOT about a specific administration action, set "
            "category to \"SKIP\" and leave event/impact empty.\n\n"
            f"Outlet: {r['outlet']}\nHeadline: {r['srcdesc']}"
        )
        try:
            msg = client.messages.create(
                model="claude-sonnet-4-6", max_tokens=400, temperature=0.2,
                messages=[{"role": "user", "content": prompt}])
            txt = "".join(b.text for b in msg.content if b.type == "text")
            txt = re.sub(r"^```json|```$", "", txt.strip()).strip()
            data = json.loads(txt)
            if data.get("category") and data["category"] != "SKIP":
                r["category"] = data.get("category", "")
                r["event"] = data.get("event", "")
                r["impact"] = data.get("impact", "")
        except Exception as ex:
            print(f"  ! draft failed for one item: {ex}", file=sys.stderr)
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--draft", action="store_true",
                    help="Use Anthropic API to pre-fill category/event/impact")
    ap.add_argument("--outdir", default="candidates")
    ap.add_argument("--pool-dir", default="",
                    help="rolling pool folder (the gather-pool branch checkout)")
    ap.add_argument("--accumulate", action="store_true",
                    help="only add today's finds to the rolling pool; no sheet")
    ap.add_argument("--tracker", default="data/Trump_Second_Term_Weekly_Tracker.xlsx",
                    help="tracker workbook, used to skip stories already logged")
    args = ap.parse_args()

    voyage_key = os.environ.get("VOYAGE_API_KEY")
    print("  Semantic matching: " + ("ON (Voyage embeddings)" if voyage_key
          else "OFF (no VOYAGE_API_KEY) — keyword matching only"))

    if args.accumulate:
        if not args.pool_dir:
            sys.exit("--accumulate needs --pool-dir")
        accumulate(args.pool_dir, voyage_key)
        return

    monday, start, end = target_week()
    print(f"Gathering week of {monday} ({start} .. {end - dt.timedelta(days=1)})")
    os.makedirs(args.outdir, exist_ok=True)
    path = os.path.join(args.outdir, f"{monday:%Y-%m-%d}.xlsx")

    new_rows, cut_rows = collect(monday, start, end, voyage_key,
                                 args.pool_dir or None, args.outdir, args.tracker)
    if args.draft:
        new_rows = ai_draft(new_rows)

    # Never clobber a sheet that's already there (you may have started
    # reviewing it, or added your own rows): keep every existing row exactly
    # as it is and add only what's new underneath.
    existing = read_sheet_rows(path) if os.path.exists(path) else []
    held = read_sheet_rows(path, MORE_TAB) if os.path.exists(path) else []
    if existing:
        print(f"  {path} already exists — keeping its {len(existing)} row(s) "
              f"and your edits, adding {len(new_rows)} new")
    rows = existing + new_rows
    in_sheet = {r.get("url") for r in rows}
    more = [r for r in held if r.get("url") not in in_sheet] + cut_rows
    write_candidates_xlsx(rows, path, more)
    print(f"Collected {len(new_rows)} new candidate item(s); wrote {path} "
          f"({len(rows)} rows)")

    # expose to the workflow (for the review issue link/name)
    out = os.environ.get("GITHUB_OUTPUT")
    if out:
        with open(out, "a") as f:
            f.write(f"file={path}\n")
            f.write(f"count={len(rows)}\n")
            f.write(f"new={len(new_rows)}\n")
            f.write(f"week={monday:%Y-%m-%d}\n")


if __name__ == "__main__":
    main()
