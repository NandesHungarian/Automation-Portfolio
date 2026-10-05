# Python Automation Scripts

Python tools for **SAP-based sales reporting and multi-year trade data analysis** in agricultural commodity trading.

All company-specific data (T-codes, user IDs, file paths, partner and location names) has been anonymized.

---

## Scripts

### `sap_sales_automation.py` — SAP Sales & Logistics Workflow Automation

**Problem solved:** The daily sales report was assembled by hand: export from SAP, convert HUF/USD prices to EUR with the day's FX rates, chase missing freight costs, reformat for management. This took 45–90 minutes every morning.

**What it does:**
- Launches SAP Logon if needed and logs in via the **SAP GUI Scripting API** (`win32com`)
- Runs the custom sales transaction for the right date range (daily report on weekdays, weekly report on Mondays)
- Exports the ALV grid to Excel and picks up the new workbook automatically
- Converts multi-currency flat prices and freight costs to EUR using the matching daily FX rate files
- Detects contracts with missing freight (non-FCA IncoTerms) and asks for the value in a **Tkinter popup**, cached per contract
- Maps columns by header name, not fixed position, so it survives SAP layout changes
- Highlights anomalies (internal partners, implausible prices, missing freight) and appends a weighted-average summary table
- Optionally calls `map_generator.py` to build the logistics map

**Time saved:** 45–90 minutes/day → a single unattended run

![Weekly summary table](../docs/images/weekly_summary.jpg)
*Summary table appended to the weekly report (test data).*

---

### `map_generator.py` — Interactive Logistics Map

Turns the finished report into a self-contained interactive HTML map (Folium / Leaflet.js):
- Routes from loading bases to delivery cities, sized by quantity and colored by commodity
- FCA volumes shown as pie-style markers at the base location
- City geocoding via OpenStreetMap Nominatim (`geopy`, no API key) with a local JSON cache
- Ambiguous city names are resolved through a small Tkinter picker
- Layer filtering per commodity via injected JavaScript

![Interactive logistics map](../docs/images/logistics_map.jpg)
*Sample output (test data).*

---

### `Historical_Trade_Data_Analyzer.py` — Multi-Year Price Normalizer

**Problem solved:** Comparing contract prices across several years was impossible without manually looking up the FX rate and pricing data valid on each contract date.

**What it does:**
- Tkinter desktop app: select the raw ERP Excel dump, processing runs in a background thread so the UI stays responsive
- Scans the archive of daily rate files and matches each contract to the rates valid on its document date (regex-based filename date parsing)
- Normalizes flat prices, truck freight and loading costs into EUR/USD
- Filters invalid rows (internal transfers, missing freight on CPT/DDP, unapproved users)
- Exports a line-by-line dataset plus a weighted-average summary per commodity to Excel

---

## Setup

Requires Windows, Microsoft Excel, and SAP GUI with scripting enabled (SAP Logon → Options → Scripting).

```bash
pip install -r requirements.txt
```

On first run `sap_sales_automation.py` creates `~/sap_config.txt` for the SAP credentials. This file is excluded by `.gitignore` and never committed. Adjust the anonymized configuration block at the top of each script (`SAP_CFG_*`, folder paths) to your environment.

> The SAP workflow also lives in its own repository with full documentation: [SAP_AgriTrade_Automation](https://github.com/NandesHungarian/SAP_AgriTrade_Automation).
