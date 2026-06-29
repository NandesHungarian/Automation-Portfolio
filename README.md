# Automation Portfolio — Nándor Magyar

> Python · Excel VBA · SAP GUI Scripting · Web Development

A collection of real-world automation tools built to eliminate repetitive manual work in **agricultural commodity trading, logistics, and business reporting**. Every script in this repository was written to solve an actual operational problem — not as a tutorial exercise.

**GitHub:** [github.com/NandesHungarian](https://github.com/NandesHungarian)
**Contact:** magyarnana97@gmail.com
**Location:** Budapest, Hungary

---

## Repository Structure

| Folder | Stack | What it does |
|--------|-------|--------------|
| [`Python/`](./Python/) | Python 3 · pandas · win32com · Tkinter | SAP integration, data analysis, interactive report generation |
| [`VBA/`](./VBA/) | Excel VBA · Outlook VBA | Excel workflow automation, email dispatch, data consolidation |
| [`Other-Projects/`](./Other-Projects/) | Astro · Tailwind CSS · Netlify | Web development — responsive, multilingual websites |

> **Note:** The [`SAP_AgriTrade_Automation`](https://github.com/NandesHungarian/SAP_AgriTrade_Automation) repository is kept **private** as it contains production-grade SAP scripting tied to real business infrastructure. Sanitized excerpts and the interactive map generator (`map_generator_.py`) are referenced below.

---

## Highlighted Projects

### 🐍 Python

**[`SAP Sales & Logistics Workflow Automation`](./Python/)**
Automatically logs into SAP ERP via GUI scripting, runs a custom transaction, exports the daily or weekly sales report, converts multi-currency prices (HUF/EUR/USD) to EUR net equivalents using daily FX rate files, flags missing freight costs with an interactive Tkinter popup, builds a formatted management summary table, and optionally generates an interactive HTML logistics map — all in a single unattended run.
`win32com` · `tkinter` · `SAP GUI Scripting` · `openpyxl`

**[`Historical Trade Data Analyzer`](./Python/Historical_Trade_Data_Analyzer.py)**
Analyses multi-year commodity trade history. Calculates period averages, detects price deviations, and exports structured Excel summaries for management review.
`pandas` · `openpyxl`

---

### ⚙️ VBA

**[`Aviso Automation`](./VBA/Aviso_Automation)** — Zero-touch daily logistics dispatch email generator. Pulls contract and shipment data from Excel, formats HTML tables, injects them into Outlook `.oft` templates, and sends to trading partners. Saves ~1 hour/day.

**[`FCA Price Indication Emailer`](./VBA/FCA_Price_Indication_Emailer.vba)** — Reads price indication tables and automatically distributes personalised offer emails to counterparties based on coverage region.

**[`Period Average Calculator`](./VBA/Period_Average_Calculator.vba)** — Computes weighted average commodity prices across configurable time windows for management reporting.

**[`Tiered Pricing & Deviation Tracker`](./VBA/Tiered_Pricing_and_Deviation_Tracker.vba)** — Flags price deviations from tiered thresholds with conditional formatting and generates a summary audit trail.

**[`Incoming Data Consolidator`](./VBA/Incoming_Data_Consolidator.vba)** — Merges and deduplicates incoming data from multiple Excel sheets into a single structured master table.

**[`Regional Coverage Report Distributor`](./VBA/Regional_Coverage_Report_Distributor.vba)** — Splits a master report into region-specific sheets and distributes each via Outlook to the responsible regional coordinator.

---

### 🌐 Web Projects

**[`Cosmetic Webpage`](./Other-Projects/)** — Responsive, multilingual (HU/EN/DE) beauty salon website built with Astro, Tailwind CSS, Netlify, and Decap CMS for client-side content management.

---

## Screenshots / Output Examples

> Sanitized screenshots and sample output files will be added here.
> Contributions and questions welcome — open an issue or reach out directly.

---

## Tech Stack Overview

![Python](https://img.shields.io/badge/Python-3776AB?style=flat&logo=python&logoColor=white)
![VBA](https://img.shields.io/badge/VBA-217346?style=flat&logo=microsoft-excel&logoColor=white)
![SAP](https://img.shields.io/badge/SAP-0FAAFF?style=flat&logo=sap&logoColor=white)
![Astro](https://img.shields.io/badge/Astro-FF5D01?style=flat&logo=astro&logoColor=white)
![Tailwind CSS](https://img.shields.io/badge/Tailwind_CSS-06B6D4?style=flat&logo=tailwind-css&logoColor=white)

---

> All company-specific data, T-codes, file paths, and user IDs in the scripts have been anonymized for public sharing.
