# Steps for Site Update

Weekly workflow for Rights Left (rightsleft.org). Updated Oct 18, 2026.

## All week, automatically

**Collect articles (rolling)** runs every 4 hours and saves relevant stories from ~25 news feeds (NPR, BBC, NYT, WaPo, NBC, CBS, ABC, PBS, Guardian, Politico, LA Times, Bloomberg, Axios, The Hill, ProPublica, The Intercept, Al Jazeera) to the `gather-pool` branch. Most feeds only hold the last day or two of stories, so this keeps a week's worth. Nothing to do here.

## Sunday, 9:00 AM Eastern

cron-job.org triggers **Gather weekly entries**. It combines the week's pool with a live pull, adds late stories from the end of last week that missed last Sunday's sheet, skips anything already offered or already on the site, writes `candidates/<date>.xlsx`, and opens a GitHub Issue.

If it doesn't show up: **Actions → Gather weekly entries → Run workflow**. Re-running is safe: it keeps your edits (including your note) and only adds new stories.

## 1. Download

Open the issue, click through to the `.xlsx`, and click the ⬇ download icon.

## 2. Review in Excel

The strongest matches are at the top (use the filter arrow on Date to sort by date instead). Pick `y` in **Include?** on every row you want; it highlights green. Category, Event and Impact are optional.

- **More matches tab** — the weaker matches that didn't fit in the top 250. It works exactly like Candidates: same Include?, Category, Event and Impact columns and dropdowns. Anything marked `y` there is ingested too.
- **Missed a story?** Add it in the first empty row under the gathered ones (either tab): `y`, headline, link, date. Outlet, Week Of and Date(s) fill themselves in at ingest. A story dated in an earlier week lands in that earlier week on the site. (There's a how-to tab in the sheet.)
- **My Note tab (optional)** — type a title and your personal note, one paragraph per row, in the yellow cells. It goes at the top of this week's newsletter under "From the Editor", signed "— Chris" (editable), and the title is added to the email subject. Leave it blank to send without a note.

Save.

## 3. Re-upload

Repo's `candidates` folder → **Add file → Upload files** → same filename → commit to `main`.

*Alternative for the note:* upload a Markdown file to `data/notes/` named for the week's Monday (e.g. `2026-10-12.md`). An optional first line `# Title`, paragraphs separated by blank lines, and a last line starting with `—` as the sign-off. A note typed in the sheet replaces this file for the same week.

## 4. Run Ingest

**Actions → Ingest approved entries → Run workflow**, file box blank.

It runs Enrich → Ingest (also saves your note to `data/notes/`) → Rebuild site → Commit → Send newsletter.

- The newsletter sends automatically, but is skipped if one already went out in the last 48 hours.
- Forgot a category? Fix the sheet, re-upload, and run Ingest again. Rows already on the site get their Category/Event/Impact updated in place (same File No.); new rows are added.
- Preview an email without sending:
  `python scripts/newsletter.py --xlsx data/Trump_Second_Term_Weekly_Tracker.xlsx --preview preview.html`

## 5. Check the site

Load rightsleft.org (no `#week=…` on the end; hard refresh if it looks stale).

## Categories (20)

Civil Rights & Minorities, Courts & SCOTUS, Dangerous Rhetoric, Democracy & Rule of Law, Economy & Tariffs, Education, Elections, Environment & Science, Executive Power, Federal Workforce, Foreign Policy & Aid, Free Speech, Healthcare, Immigration, Infrastructure and History, Law Enforcement, LGBTQ+ Rights, National Security, Press Freedom, Public Health.
