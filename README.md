# Automation Portfolio — Nándor Magyar

> Python · Excel VBA · SAP GUI Scripting · Web Development

A collection of real-world automation tools built to eliminate repetitive manual work in logistics, agricultural trading, and business reporting. Every script in this repository was written to solve an actual problem I encountered on the job — not as a tutorial exercise.

**GitHub:** [github.com/NandesHungarian](https://github.com/NandesHungarian)  
**Contact:** magyarnana97@gmail.com  
**Location:** Budapest, Hungary

---

## Repository Structure

| Folder | Stack | What it does |
|--------|-------|--------------|
| [`Python/`](./Python/) | Python 3, pandas, win32com, Tkinter | Data analysis, SAP integration, interactive report generation |
| [`VBA/`](./VBA/) | Excel VBA, Outlook VBA | Excel-based workflow automation, email dispatch, data consolidation |
| [`Other-Projects/`](./Other-Projects/) | Astro, Tailwind CSS, Netlify | Web development projects — responsive, multilingual websites |

---

## Highlighted Projects

### 🐍 Python

**[`SAP Sales & Logistics Workflow Automation`](./Python/sap_sales_automation_.py)**  
Automatically logs into SAP via GUI scripting, exports the daily/weekly sales report, calculates EUR-equivalent net prices from multi-currency data (HUF/EUR/USD), flags missing freight rates with a Tkinter popup, and writes a formatted Excel management report — all in one unattended run.  
`win32com` · `tkinter` · `SAP GUI Scripting` · `openpyxl`

**[`Historical Trade Data Analyzer`](./Python/Historical_Trade_Data_Analyzer.py)**  
Analyses multi-year commodity trade history. Calculates period averages, detects price deviations, and exports structured Excel summaries for management review.  
`pandas` · `openpyxl`

---

### ⚙️ VBA

**[`Aviso Automation`](./VBA/Aviso_Automation)** — Zero-touch daily logistics dispatch email generator. Pulls data from Excel, formats HTML tables, injects them into Outlook `.oft` templates and sends to partners. Saves ~1 hour/day.  

**[`FCA Price Indication Emailer`](./VBA/FCA_Price_Indication_Emailer.vba)** — Reads price indication tables and automatically distributes personalised emails to counterparties based on coverage region.  

**[`Period Average Calculator`](./VBA/Period_Average_Calculator.vba)** — Computes weighted average commodity prices across configurable time windows for management reporting.  

**[`Tiered Pricing & Deviation Tracker`](./VBA/Tiered_Pricing_and_Deviation_Tracker.vba)** — Flags price deviations from tiered thresholds with conditional formatting and generates a summary audit trail.  

**[`Incoming Data Consolidator`](./VBA/Incoming_Data_Consolidator.vba)** — Merges and deduplicates incoming data from multiple Excel sheets into a single structured master table.  

**[`Regional Coverage Report Distributor`](./VBA/Regional_Coverage_Report_Distributor.vba)** — Splits a master report into region-specific sheets and distributes each via Outlook to the responsible regional coordinator.

---

### 🌐 Web Projects

**[`Cosmetic Webpage`](./Other-Projects/)** — Responsive, multilingual (HU/EN/DE) beauty salon website built with Astro, Tailwind CSS, Netlify, and Decap CMS for client-side content management.

---

## Screenshots / Output Examples

> *Sanitized screenshots and sample outputs will be added here. Contributions welcome.*

---

## Tech Stack Overview

![Python](https://img.shields.io/badge/Python-3776AB?style=flat&logo=python&logoColor=white)
![VBA](https://img.shields.io/badge/VBA-217346?style=flat&logo=microsoft-excel&logoColor=white)
![SAP](https://img.shields.io/badge/SAP-0FAAFF?style=flat&logo=sap&logoColor=white)
![Astro](https://img.shields.io/badge/Astro-FF5D01?style=flat&logo=astro&logoColor=white)
![Tailwind CSS](https://img.shields.io/badge/Tailwind_CSS-06B6D4?style=flat&logo=tailwind-css&logoColor=white)

---

> All company-specific data, T-codes, file paths, and user IDs in the Python scripts have been anonymized for public sharing.
