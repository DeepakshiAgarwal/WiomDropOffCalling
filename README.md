# Wiom APP Drop-off Calling — Drop-off Recovery Desk

A one-page calling dashboard for APP drop-off leads: funnel, intent breakdown,
top reasons, and decision-day timing. Every number is computed with the same
formula logic as the **Executive Summary - APP Dropoff** tab of
`Wiom_Dropoff_Calling_Pack_03Sep2026.xlsx` (COUNTUNIQUE / COUNTIFS for volumes,
MEDIAN / AVERAGE / PERCENTILE for decision-day timing), validated line-for-line
against that tab's own cached values.

## Files

| File | What it is |
|---|---|
| `index.html` / `dropoff_dashboard.html` | The rendered dashboard (identical content, two names — `index.html` so GitHub Pages serves it at the repo root). |
| `template.html.j2` | Jinja2 template the dashboard is rendered from. |
| `build_dashboard.py` | Rebuilds the dashboard from a fresh export of the source workbook. |

## Rebuilding from fresh data

1. Open the source Google Sheet and download the workbook as `.xlsx`
   (File → search "Download" → **Microsoft Excel (.xlsx)**).
2. Run:
   ```bash
   pip install openpyxl pandas jinja2
   python build_dashboard.py "<path-to-downloaded-xlsx>" "dropoff_dashboard.html"
   cp dropoff_dashboard.html index.html
   ```
3. Commit and push.

The script reads two sheets from the workbook:
- `CALL SHEET - APP Drop offs` — raw call-level data (funnel, intent bifurcation, decision-day gap).
- `DAILY SUMMARY -APP Drop offs ` — a reason-code → friendly-label → outcome-bucket lookup table (rows 70-99), used for the Top-5-reasons and full-breakdown sections.

## Viewing it live

Enable **Settings → Pages → Deploy from branch → `main` / `/ (root)`** on this
repo to get a public URL serving `index.html`.

## Data note

This is a **static snapshot**, not a live connection to the Google Sheet — it
reflects the data as of whenever `build_dashboard.py` was last run, not the
current live sheet. Rebuild and push again to refresh.
