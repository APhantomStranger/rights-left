# Rights Left — automated weekly pipeline (Gather → Approve → Ingest)

This repo keeps the tracker workbook up to date every week with a human in the loop.

- **Gather** (automatic, Sundays): collects the past week's political news from RSS
  feeds into a review file — `candidates/<monday>.csv`.
- **Approve** (you, ~5–10 min): edit that CSV — fill in category/event/impact and
  mark the rows to keep with `include = y`.
- **Ingest** (one click): appends your approved rows into
  `data/Trump_Second_Term_Weekly_Tracker.xlsx` in the exact existing format
  (fonts, banding, hyperlinks, File-No. sequence, and Summary totals all match).

Nothing is published without your approval, and the ingest step never invents
content — it only formats and files what you approved.

```
rights-left/
├─ data/Trump_Second_Term_Weekly_Tracker.xlsx   ← the master workbook
├─ scripts/gather.py                             ← collects candidates
├─ scripts/ingest.py                             ← appends approved rows
├─ candidates/                                   ← weekly review CSVs land here
│  └─ processed/                                 ← ingested CSVs are moved here
├─ requirements.txt
└─ .github/workflows/
   ├─ gather.yml     (schedule: Sundays 13:00 UTC + manual)
   └─ ingest.yml     (manual "Run workflow" button)
```

---

## One-time setup

1. **Create the repository.**
   - On GitHub: **New repository** → name it `rights-left` → **Private** is fine → Create.
   - Upload these files preserving the folder structure. Easiest path:
     install [Git](https://git-scm.com/downloads), then in a terminal:
     ```bash
     git clone https://github.com/<you>/rights-left.git
     # copy the files from this bundle into the cloned folder, then:
     cd rights-left
     git add .
     git commit -m "Initial pipeline"
     git push
     ```
     (Or use the GitHub website's **Add file → Upload files**, but you must create
     the `scripts/`, `candidates/`, `data/`, and `.github/workflows/` folders by
     typing the path when uploading, e.g. `scripts/gather.py`.)

2. **Allow Actions to write to the repo.**
   - **Settings → Actions → General → Workflow permissions** →
     select **Read and write permissions** → **Save**.

3. **(Optional) Turn on AI pre-drafting.** Skip this to keep the pipeline free and
   fully manual; the collector still works, you just write category/event/impact
   yourself during approval.
   - Get an API key at console.anthropic.com (this is a **paid** product, billed
     per use — pennies per week at this volume).
   - **Settings → Secrets and variables → Actions → New repository secret**
     → Name: `ANTHROPIC_API_KEY` → paste the key → **Add secret**.
   - In `.github/workflows/gather.yml`, change the collect step to:
     `run: python scripts/gather.py --draft`

That's it. The Sunday schedule is now live.

---

## The weekly routine

1. **Sunday** — the **Gather** workflow runs on its own. It commits
   `candidates/<monday>.csv` and opens a GitHub **Issue** titled
   "Review candidates — week of …". (Watch the **Actions** tab or your email.)

2. **Approve** — open the CSV (link is in the issue). Click the **pencil/Edit**
   icon and, for each item worth keeping:
   - set **category** to one of the 17 the workbook uses (listed at the top of
     `gather.py`),
   - write a short **event** (what the administration did) and **impact**
     (why critics/courts/data call it harmful),
   - set **include** to `y`.
   Delete or leave blank the rows you don't want. **Commit** to `main`.
   *(With `--draft` on, category/event/impact arrive pre-filled — you just check
   them and set `include`.)*

3. **Ingest** — go to **Actions → Ingest approved entries → Run workflow**
   (leave the file box blank to use the newest CSV) → **Run**. In under a minute
   the workbook in `data/` is updated and committed, and the processed CSV is
   moved to `candidates/processed/`. Download the updated `.xlsx` from the repo.

To test before Sunday, run **Gather** manually the same way (it has a
**Run workflow** button too).

---

## Good to know (honest caveats)

- **Cron is UTC.** `0 13 * * 0` = Sundays 13:00 UTC. Adjust the number for your
  timezone. Scheduled runs can be delayed a few minutes under GitHub load.
- **60-day idle rule.** GitHub disables scheduled workflows in a repo with no
  activity for 60 days. Since ingest commits every week, normal use keeps it
  alive; if you ever pause, re-enable it in the Actions tab.
- **Feeds change.** Outlets occasionally move or retire RSS URLs. A feed that
  errors is skipped, not fatal. Edit the `FEEDS` dict in `gather.py` to add or
  swap sources. Reuters dropped public RSS, so it isn't included by default.
- **The collector is a net, not a judge.** It over-collects on keywords so you
  don't miss things; pruning is the whole point of the approve step. It only sees
  what the feeds carry — a story no feed surfaces won't appear.
- **Duplicates are handled.** Ingest skips any row whose source URL is already in
  the workbook, so re-running or overlapping weeks won't double-file.
- **Want the website updated too?** Ask Claude for an xlsx-driven `build_site.py`
  and uncomment the "Rebuild website" step in `ingest.yml`; each ingest will then
  regenerate `rights-left.html` from the workbook.
