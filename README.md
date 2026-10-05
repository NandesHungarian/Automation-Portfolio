# Automation Portfolio — Nándor Magyar

> Python · Excel VBA · SAP GUI Scripting · Outlook automation

Real-world automation tools built to eliminate repetitive manual work in **agricultural commodity trading, logistics, and business reporting**. Every script here was written to solve an actual operational problem at work, not as a tutorial exercise.

**Contact:** magyarnana97@gmail.com &nbsp;·&nbsp; **Location:** Malta

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
Logs into SAP ERP via GUI scripting, runs a custom transaction, exports the daily or weekly sales report, converts HUF/USD prices to EUR net equivalents using daily FX rate files, asks for missing freight costs in a Tkinter popup, builds a formatted management summary, and optionally generates an interactive HTML logistics map. Replaces a 45–60 minute manual process every morning.
`win32com` · `SAP GUI Scripting` · `tkinter` · `folium` · `geopy`
→ Full documentation: [SAP_AgriTrade_Automation](https://github.com/NandesHungarian/SAP_AgriTrade_Automation)

![Interactive logistics map](docs/images/logistics_map.jpg)
*Generated logistics map: bases, delivery routes and product mix (test data).*

![Weekly summary table](docs/images/weekly_summary.jpg)
*Weekly summary appended to the report: quantity and weighted average net EUR price (test data).*

**Historical Trade Data Analyzer** — [`Historical_Trade_Data_Analyzer.py`](./Python/Historical_Trade_Data_Analyzer.py)
Desktop app for multi-year analysis. It opens hundreds of daily rate files, matches every contract to the forward rates valid on its date, normalizes all prices to EUR/USD, filters invalid rows, and exports weighted-average summaries for management review.
`pandas` · `numpy` · `tkinter` · `threading`

---

### ⚙️ VBA

| Macro | What it does | Time saved |
|---|---|---|
| [`Aviso_Automation`](./VBA/Aviso_Automation.vba) | Prepares a daily email for each partner with what they collected and what is still waiting, using their own Outlook template | ~1 hour/day |
| [`FCA_Price_Indication_Emailer`](./VBA/FCA_Price_Indication_Emailer.vba) | Sends the daily price table as an image in an Outlook email and saves the prices to a log sheet, building a price history for later analysis | |
| [`Email_Attachment_Importer`](./VBA/Email_Attachment_Importer.vba) | One click finds today's report email in Outlook and pastes its Excel attachment into the right sheet. Easy to reuse for other daily emails | |
| [`Period_Average_Calculator`](./VBA/Period_Average_Calculator.vba) | Calculates soybean meal price averages for the same periods as sunflower and rapeseed meal, even though the two files store dates differently | |
| [`Tiered_Pricing_and_Deviation_Tracker`](./VBA/Tiered_Pricing_and_Deviation_Tracker.vba) | Sends daily prices to three partner groups with their own price tables pasted into the email, and logs each period's price against the reference price | |
| [`Incoming_Data_Consolidator`](./VBA/Incoming_Data_Consolidator.vba) | Merges and deduplicates regional data files into one master table | |
| [`Regional_Coverage_Report_Distributor`](./VBA/Regional_Coverage_Report_Distributor.vba) | Splits the master report per coordinator and emails each slice | ~45 min/week |

Details for each macro: [`VBA/README.md`](./VBA/README.md)

---

> All company-specific data — partner names, T-codes, user IDs, file paths and email addresses — has been anonymized for public sharing.
