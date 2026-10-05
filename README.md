# Automation Portfolio — Nándor Magyar

> Python · Excel VBA · SAP GUI Scripting · Outlook automation

Real-world automation tools built to eliminate repetitive manual work in **agricultural commodity trading, logistics, and business reporting**. Every script here was written to solve an actual operational problem at work, not as a tutorial exercise.

**Contact:** magyarnana97@gmail.com &nbsp;·&nbsp; **Location:** Budapest, Hungary

---

## Repository Structure

| Folder | Stack | What it does |
|---|---|---|
| [`Python/`](./Python) | Python 3 · pandas · win32com · Tkinter · Folium | SAP integration, FX normalization, data analysis, interactive maps |
| [`VBA/`](./VBA) | Excel VBA · Outlook VBA | Email dispatch, data import and consolidation, pricing checks |

---

## Highlighted Projects

### 🐍 Python

**SAP Sales & Logistics Workflow Automation** — [`sap_sales_automation.py`](./Python/sap_sales_automation.py) · [`map_generator.py`](./Python/map_generator.py)
Logs into SAP ERP via GUI scripting, runs a custom transaction, exports the daily or weekly sales report, converts HUF/USD prices to EUR net equivalents using daily FX rate files, asks for missing freight costs in a Tkinter popup, builds a formatted management summary, and optionally generates an interactive HTML logistics map. Replaces a 45–90 minute manual process every morning.
`win32com` · `SAP GUI Scripting` · `tkinter` · `folium` · `geopy`
→ Full documentation: [SAP_AgriTrade_Automation](https://github.com/NandesHungarian/SAP_AgriTrade_Automation)

![Interactive logistics map](docs/images/logistics_map.jpg)
*Generated logistics map: bases, delivery routes and product mix (test data).*

![Weekly summary table](docs/images/weekly_summary.jpg)
*Weekly summary appended to the report: quantity and weighted average net EUR price (test data).*

**Historical Trade Data Analyzer** — [`Historical_Trade_Data_Analyzer.py`](./Python/Historical_Trade_Data_Analyzer.py)
Desktop app that matches multi-year contract history to the FX and pricing data valid on each contract date, normalizes all prices to EUR/USD, filters invalid rows, and exports weighted-average summaries for management review.
`pandas` · `numpy` · `tkinter` · `threading`

---

### ⚙️ VBA

| Macro | What it does | Time saved |
|---|---|---|
| [`Aviso_Automation`](./VBA/Aviso_Automation.vba) | Builds daily logistics dispatch emails from Excel data into Outlook `.oft` templates, per partner | ~1 hour/day |
| [`FCA_Price_Indication_Emailer`](./VBA/FCA_Price_Indication_Emailer.vba) | Sends personalised FCA price offers to counterparties based on coverage region | |
| [`Email_Attachment_Importer`](./VBA/Email_Attachment_Importer.vba) | Finds today's report email in Outlook and imports its Excel attachment into the workbook | |
| [`Period_Average_Calculator`](./VBA/Period_Average_Calculator.vba) | Quantity-weighted average prices over configurable periods | |
| [`Tiered_Pricing_and_Deviation_Tracker`](./VBA/Tiered_Pricing_and_Deviation_Tracker.vba) | Flags contracts booked outside the approved pricing grid, with an audit sheet | |
| [`Incoming_Data_Consolidator`](./VBA/Incoming_Data_Consolidator.vba) | Merges and deduplicates regional data files into one master table | |
| [`Regional_Coverage_Report_Distributor`](./VBA/Regional_Coverage_Report_Distributor.vba) | Splits the master report per coordinator and emails each slice | ~45 min/week |

Details for each macro: [`VBA/README.md`](./VBA/README.md)

---

> All company-specific data — partner names, T-codes, user IDs, file paths and email addresses — has been anonymized for public sharing.
