name: Ingest approved entries

on:
  workflow_dispatch:
    inputs:
      file:
        description: 'Approved CSV path (leave blank to use the newest in candidates/)'
        required: false
        default: ''

permissions:
  contents: write            # to commit the updated workbook

jobs:
  ingest:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4

      - uses: actions/setup-python@v5
        with:
          python-version: '3.12'

      - name: Install dependencies
        run: pip install -r requirements.txt

      - name: Enrich approved entries from their source articles
        env:
          # Needs this secret set (Settings → Secrets and variables → Actions).
          # Without it, this step prints a notice and does nothing — the rest
          # of the pipeline still works, you just fill fields in by hand.
          ANTHROPIC_API_KEY: ${{ secrets.ANTHROPIC_API_KEY }}
        run: |
          python scripts/enrich.py \
            ${{ inputs.file != '' && format('--csv {0}', inputs.file) || '' }}

      - name: Install LibreOffice (recalculates formula values)
        run: sudo apt-get update && sudo apt-get install -y libreoffice-calc

      - name: Append approved rows into the workbook
        run: |
          python scripts/ingest.py \
            --xlsx "data/Trump_Second_Term_Weekly_Tracker.xlsx" \
            ${{ inputs.file != '' && format('--csv {0}', inputs.file) || '' }} \
            --recalc

      - name: Rebuild website from updated workbook
        run: python scripts/build_site.py --xlsx "data/Trump_Second_Term_Weekly_Tracker.xlsx"

      - name: Commit updated workbook
        run: |
          git config user.name  "rights-left-bot"
          git config user.email "actions@users.noreply.github.com"
          git add data/ candidates/ index.html
          git commit -m "Ingest approved entries ($(date -u +%Y-%m-%d))" || echo "nothing to commit"
          git push
