# Timesheet Converter

A static web page that turns a Jibble **Monthly Raw Time Entries** CSV into the company's
`FORMAT Timesheet.xlsx` layout. It runs entirely in the browser, with Python (openpyxl) running
in [Pyodide](https://pyodide.org). Files are never uploaded anywhere.

## Monthly use

1. In Jibble: Reports → **Monthly Raw Time Entries** → export CSV.
2. Open the site, drop in the CSV, click **Export timesheet**. The .xlsx downloads.

Under the button, the page lists the days left blank (no clock-out) and shifts over 14 hours.

What the page does not do (do it in Excel afterwards if needed):
- **Sakit / izin / cuti:** type `S`, `I` or `C` in the person's Status row in MASTER DATA. The totals
  formulas count them.
- **Jabatan / Status Karyawan:** the Jabatan cell shows "Proyek : <group>"; add the role above it.

The group order and who is listed first come from `py/config.py` (project names only, no people).
New groups go after the listed ones, A to Z. People within a group are sorted by name.

`tools/cli.py` still supports the extras for local use: `--prev` (last month's file, so roles carry
over through its hidden `_ROSTER` sheet) and `--manual` (a CSV of sakit/izin/cuti days).

## Status rules

Applied in this order, per person per day:

| Situation | Status |
|---|---|
| IN or OUT missing | blank (never an automatic `a`) |
| Sunday (Libur) with attendance | `L` |
| Lokasi contains WFH (Jibble "WFA" is WFH) | `W` |
| OUT − IN ≥ 6 h | `1` |
| OUT − IN < 6 h | `0.5` |
| Entered by hand | `S`, `I`, `C` (or `a`) |

How Jibble punches are read (`py/parse.py`):

- Punch time is **Dibuat pada**, the creation time, which has seconds. An entry a manager adds later
  counts at the moment it was added (`tools/cli.py --use-waktu` uses Tanggal + Waktu instead).
- An In while already clocked in is an activity switch, and the day's Lokasi lists the activities
  joined by `> `. StartBreak ends a session.
- An overnight shift counts on the day it started.
- An Out followed within 5 minutes by an In the next morning is a forgotten clock-out. That
  day's OUT is missing, and so is the OUT of any session longer than 24 h.
- The day's OUT is the last valid clock-out of that day.

## Privacy

GitHub Pages sites and their repos are public on the free plan. This repo holds **code only**:

- `.gitignore` excludes every `*.xlsx` and `*.csv` except `template/format-timesheet-clean.xlsx`,
  which has no people, notes or images in it.
- The deploy workflow fails if any other spreadsheet or CSV gets committed.
- The roster lives inside each month's output workbook (the hidden `_ROSTER` sheet), not in the repo.

## Layout

```
index.html, app.js, style.css   the page
worker.js                       loads Pyodide + vendored openpyxl and runs py/ off the main thread
py/config.py                    default group order (no personal data)
py/parse.py                     CSV → one IN / OUT / Lokasi per person per day
py/roster.py                    _ROSTER sheet: read, seed from MASTER DATA, reconcile, write
py/build.py                     fills the template (MASTER DATA, 01 REKAP ABSEN, logs)
py/api.py                       JSON entry points used by the worker and the CLI
template/                       the clean template (made with tools/make_clean_template.py)
vendor/                         pinned openpyxl + et_xmlfile wheels (not part of Pyodide)
tools/cli.py                    run the same converter locally
tests/test_rules.py             rule tests on made-up data (run in CI)
```

## Local use and development

```bash
pip install vendor/et_xmlfile-2.0.0-py3-none-any.whl vendor/openpyxl-3.1.5-py2.py3-none-any.whl
python tests/test_rules.py
python tools/cli.py export.csv --prev "last month.xlsx" --manual sakit-izin.csv --out "Timesheet.xlsx"
python -m http.server 8765        # then open http://localhost:8765
```

Keep real exports in `local/`, which git ignores.

## Deploying

Push to `main` on GitHub, then go to Settings → Pages → Source: **GitHub Actions**. The workflow in
`.github/workflows/pages.yml` runs the tests and publishes the site.
