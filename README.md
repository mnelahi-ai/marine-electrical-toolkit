# Marine Electrical Toolkit

Engineering tools for marine electrical design, document control and load-list
production. Built around one principle: **no guessing — every value cites a
source, every unknown is flagged.**

## Tools

| Tool | What it does | Run |
|---|---|---|
| `hull_ingest.py` | Reads EVERY file in a hull folder (PDF/XLSX/CSV/TXT), extracts candidate kW/V/Phase/Hz data with source-document citation → `Master_Source_List_<hull>.xlsx`. Scanned PDFs flagged `[MISSING-TEXT]`. | `python hull_ingest.py "H675"` |
| `workbook_builder.py` | Builds a Load List + Alarm List workbook skeleton (Cover / 11-section Load List / 9-group Alarm List, PF 0.85 formula block) from a JSON config. Unknowns written as `[TBC]`, never guessed. | `python workbook_builder.py hull_config.json` |
| `cable_schedule.py` | Generates a Cable Schedule from any Load List workbook — current per row, indicative conductor size (IEC 60092-352-style table), lengths `[TBC]`. | `python cable_schedule.py "H675_Load List...xlsx"` |
| `research_queue.py` | Scans workbooks for `[TBC]` / `[MISSING]` / blank kW/Current/Remarks → `Research_Queue.xlsx` with ready-made search queries. | `python research_queue.py *.xlsx` |
| `folder_organiser.py` | SHA256 dedupe report, backup archiving (move-only, never delete), folder auto-index. Dry-run by default. | `python folder_organiser.py . --index` |
| `webapp/index.html` | 11 free marine electrical calculators (short circuit, switchboard rating, fuel consumption, motor earth fault, transformer/motor FLC, volt drop, panel heat loss, earth cable sizing, energy cost, Ex e enclosure heat dissipation). Single file, no backend. | open in browser |
| `webapp/checklists.html` | 12 interactive commissioning checklists (HV swgr, LV switchboard, transformer, VSD, genset, UPS/DDUPS, busway, cabling, power turn-on, documentation, FAT MSB, Ex-proof DB FAT) with progress tracking, CSV export and print. | open in browser |
| `webapp/testforms.html` | Continuity & insulation resistance (megger) test sheet with auto pass/fail + CSV export, and a printable General Test Certificate. | open in browser |
| `webapp/reference.html` | Reference tables: ANSI protection functions, motor FLC/trip sizes (computed), XLPE cable parameters, IR test voltages, typical lux levels, Ex marking guide. | open in browser |

## Install

```
pip install openpyxl pdfminer.six
```

## Publish the web app on GitHub Pages (free)

1. Create a repo (e.g. `marine-electrical-toolkit`) at github.com.
2. Upload this folder's contents (`webapp/index.html` at minimum).
3. Repo Settings → Pages → Source: `main` branch → root (or `/webapp`).
4. Your site goes live at `https://<username>.github.io/marine-electrical-toolkit/`.

Monetisation options once traffic exists: Google AdSense, "buy me a coffee"
link, affiliate links to test-equipment suppliers, or paid downloads of the
genericized checklist/FAT template pack (strip all project numbers and client
names from any document before publishing — only publish concepts you own).

## Disclaimers

All calculators and generated schedules are **indicative engineering aids**.
Results must be verified by a qualified engineer against the applicable class
rules (LR/BV/DNV/ABS, NSCV/AMSA, etc.), flag requirements and manufacturer
data before use in design, construction or survey submissions.
