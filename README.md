# Automation Portfolio — Nándor Magyar

> Python · Excel VBA · SAP GUI Scripting · Web Development

A collection of real-world automation tools built to eliminate repetitive manual work in **agricultural commodity trading, logistics, and business reporting**. Every script in this repository was written to solve an actual operational problem — not as a tutorial exercise.

**GitHub:** github.com/NandesHungarian &nbsp;·&nbsp; **Contact:** magyarnana97@gmail.com &nbsp;·&nbsp; **Location:** Budapest, Hungary

---

## Repository Structure

| Folder | Stack | What it does |
|---|---|---|
| `Python/` | Python 3 · pandas · win32com · Tkinter | SAP integration, data analysis, interactive report generation |
| `VBA/` | Excel VBA · Outlook VBA | Excel workflow automation, email dispatch, data consolidation |
| `Web/` | Astro · Tailwind CSS · Netlify | Responsive, multilingual websites for real clients |

> **Note:** The full `SAP_AgriTrade_Automation` repository is kept **private** as it contains production-grade SAP scripting tied to real business infrastructure. A sanitized, fully functional excerpt (`sap_sales_automation_.py`) is available in the `Python/` folder.

---

## Highlighted Projects

### 🐍 Python

**`SAP Sales & Logistics Workflow Automation`** [`sap_sales_automation_.py`](./Python/sap_sales_automation_.py)
Automatically logs into SAP ERP via GUI scripting, runs a custom transaction, exports the daily or weekly sales report, converts multi-currency prices (HUF/EUR/USD) to EUR net equivalents using daily FX rate files, flags missing freight costs with an interactive Tkinter popup, builds a formatted management summary table, and optionally generates an interactive HTML logistics map — all in a single unattended run.
`win32com` · `tkinter` · `SAP GUI Scripting` · `openpyxl`

**`Historical Trade Data Analyzer`** [`Historical_Trade_Data_Analyzer.py`](./Historical_Trade_Data_Analyzer.py)
Analyses multi-year commodity trade history. Calculates period averages, detects price deviations, and exports structured Excel summaries for management review.
`pandas` · `openpyxl`

---

### ⚙️ VBA

**`Aviso Automation`** — Zero-touch daily logistics dispatch email generator. Pulls contract and shipment data from Excel, formats HTML tables, injects them into Outlook `.oft` templates, and sends to trading partners. Saves ~1 hour/day.

**`FCA Price Indication Emailer`** — Reads price indication tables and automatically distributes personalised offer emails to counterparties based on coverage region.

**`Period Average Calculator`** — Computes weighted average commodity prices across configurable time windows for management reporting.

**`Tiered Pricing & Deviation Tracker`** — Flags price deviations from tiered thresholds with conditional formatting and generates a summary audit trail.

**`Incoming Data Consolidator`** — Merges and deduplicates incoming data from multiple Excel sheets into a single structured master table.

**`Regional Coverage Report Distributor`** — Splits a master report into region-specific sheets and distributes each via Outlook to the responsible regional coordinator.

---

### 🌐 Web

**[`Cosmetic-webpage`](https://github.com/NandesHungarian/Cosmetic-webpage)** — Responsive, multilingual (HU/EN/DE) beauty salon website built with Astro, Tailwind CSS, Netlify, and Decap CMS for client-side content management.

**[`italiano-b2`](https://github.com/NandesHungarian/italiano-b2)** — Modern responsive website built with Astro and Tailwind CSS.

---

## Screenshots / Output Examples

> Sanitized screenshots and sample output files will be added here.
> Contributions and questions welcome — open an issue or reach out directly.

---

## Tech Stack Overview

> All company-specific data, T-codes, file paths, and user IDs in the scripts have been anonymized for public sharing.
