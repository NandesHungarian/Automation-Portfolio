# Automation Portfolio — Nándor Magyar

> Python · Excel VBA · SAP GUI Scripting · Web Development

A collection of real-world automation tools built to eliminate repetitive manual work in logistics, agricultural commodity trading, and business reporting. Every script in this repository was written to solve an actual problem encountered on the job — not as a tutorial exercise.

**GitHub:** [github.com/NandesHungarian](https://github.com/NandesHungarian)  
**Contact:** magyarnana97@gmail.com  
**Location:** Budapest, Hungary

> All company-specific data, T-codes, file paths, and user IDs in the Python scripts have been anonymized for public sharing.

---

## Repository Structure

| Folder | Stack | What it does |
|--------|-------|--------------|
| [`Python/`](./Python/) | Python 3, pandas, win32com, tkinter, openpyxl | SAP GUI integration, trade data analysis, interactive report generation |
| [`VBA/`](./VBA/) | Excel VBA, Outlook VBA | Excel-based workflow automation, email dispatch, data consolidation |
| [`Other-Projects/`](./Other-Projects/) | Astro, Tailwind CSS, Netlify, HTML/CSS | Web development projects and miscellaneous builds |

---

## Python Projects

### [`SAP Sales & Logistics Workflow Automation`](./Python/)

Fully unattended end-to-end pipeline that handles the entire daily/weekly logistics reporting cycle:

1. **SAP login & report extraction** — opens SAP GUI via `win32com`, authenticates, navigates to a custom transaction, sets the date range automatically (daily or weekly depending on the day of the week), and exports the ALV report to Excel
2. **Multi-currency price conversion** — reads daily EUR/HUF and USD/HUF base price files, matches exchange rates to each contract’s document date, and calculates EUR-equivalent net flat prices for all rows
3. **Missing freight detection** — identifies contracts where IncoTerms require a freight rate (DDP, CPT, etc.) but none is present; prompts the user via a `tkinter` popup to enter the missing rate on the spot, caches it, and backfills all related rows automatically
4. **Report formatting** — reorders columns to a fixed layout, applies number formatting, highlights problematic rows in red, adds autofilter for valid sales users, and auto-fits all columns
5. **Summary table generation** — builds a pivot-style summary at the bottom of the sheet grouping quantity and weighted average net price by location, commodity, and crop year; weekly mode adds a Friday-only breakdown
6. **Interactive map prompt** — after saving, offers to launch `map_generator_.py` to produce an HTML logistics map from the finished report

**Tech:** `win32com` · `tkinter` · `openpyxl` · `pandas` · SAP GUI Scripting API

---

### [`Historical Trade Data Analyzer`](./Python/Historical_Trade_Data_Analyzer.py)

Analyses multi-year commodity trade history exported from internal systems. Calculates period averages, detects price deviations from baseline, and exports structured Excel summaries formatted for management review.

**Tech:** `pandas` · `openpyxl`

---

### [`Logistics Map Generator`](./Python/) *(companion to SAP automation)*

Generates an interactive HTML map from the processed logistics report. Plots delivery routes, partner locations, and contract volumes. Launched automatically at the end of the SAP automation run or manually from the same Excel output.

**Tech:** `folium` / HTML · `pandas` · `openpyxl`

---

## VBA Projects

### [`Aviso Automation`](./VBA/Aviso_Automation)

Zero-touch daily logistics dispatch email generator. Reads contract and delivery data from Excel, builds formatted HTML tables, injects them into Outlook `.oft` email templates, and sends personalised notifications to logistics partners. Saves approximately 1 hour per day of manual copy-paste work.

**Tech:** Excel VBA · Outlook VBA · HTML email templating

---

### [`FCA Price Indication Emailer`](./VBA/FCA_Price_Indication_Emailer.vba)

Reads FCA price indication tables from a master Excel sheet and automatically distributes personalised price emails to counterparties, filtered by their regional coverage. One click replaces individually addressed emails to each partner.

**Tech:** Excel VBA · Outlook VBA

---

### [`Period Average Calculator`](./VBA/Period_Average_Calculator.vba)

Computes weighted average commodity prices across configurable time windows (daily, weekly, monthly, crop-year). Used for management reporting and tender price benchmarking.

**Tech:** Excel VBA

---

### [`Tiered Pricing & Deviation Tracker`](./VBA/Tiered_Pricing_and_Deviation_Tracker.vba)

Flags price deviations from tiered pricing thresholds using conditional formatting. Generates a summary audit trail showing which contracts exceeded acceptable deviation bands and by how much.

**Tech:** Excel VBA

---

### [`Incoming Data Consolidator`](./VBA/Incoming_Data_Consolidator.vba)

Merges and deduplicates incoming data received from multiple Excel sheets (different partners, different formats) into a single normalised master table. Handles varying column orders and missing fields.

**Tech:** Excel VBA

---

### [`Regional Coverage Report Distributor`](./VBA/Regional_Coverage_Report_Distributor.vba)

Splits a master coverage report into region-specific sub-reports and distributes each automatically via Outlook to the responsible regional coordinator. Eliminates the manual filter-copy-send cycle.

**Tech:** Excel VBA · Outlook VBA

---

## Other Projects

See [`Other-Projects/`](./Other-Projects/) for web development work and miscellaneous builds.

---

## Screenshots & Sample Outputs

> Sanitized screenshots and sample output files will be added here. If you’d like to see a demo of any script, feel free to reach out.

---

## Tech Stack

![Python](https://img.shields.io/badge/Python-3776AB?style=flat&logo=python&logoColor=white)
![VBA](https://img.shields.io/badge/Excel_VBA-217346?style=flat&logo=microsoft-excel&logoColor=white)
![SAP](https://img.shields.io/badge/SAP_GUI_Scripting-0FAAFF?style=flat&logo=sap&logoColor=white)
![Astro](https://img.shields.io/badge/Astro-FF5D01?style=flat&logo=astro&logoColor=white)
![Tailwind CSS](https://img.shields.io/badge/Tailwind_CSS-06B6D4?style=flat&logo=tailwind-css&logoColor=white)
![Netlify](https://img.shields.io/badge/Netlify-00C7B7?style=flat&logo=netlify&logoColor=white)
