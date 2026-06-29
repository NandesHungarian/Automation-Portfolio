# Python Automation Scripts

Real-world Python scripts built for **agricultural commodity trading and logistics operations**. Each script solves a specific, recurring operational problem that previously required manual effort.

---

## Scripts

### `sap_sales_automation_.py` — SAP Sales & Logistics Workflow Automation

**Problem solved:** Generating the daily/weekly sales management report required logging into SAP manually, running a custom transaction, exporting to Excel, then spending 30–60 minutes on manual price calculations, currency conversions, and formatting. This script does all of it automatically.

**What it does:**
1. Launches SAP GUI and logs in automatically from a local credentials file
2. Navigates to the custom sales transaction, sets the date range (daily on weekdays, weekly on Monday), and exports the report to Excel
3. Detects missing freight costs for non-FCA contracts and prompts the user with a Tkinter popup — caches the entered value for the entire contract
4. Loads daily FX rate files (EUR/HUF, USD/HUF) and converts all Flat Prices and freight costs to EUR net equivalents
5. Reorders columns to a fixed management layout, applies conditional formatting (red rows for anomalies), and adds a pivot-style summary table at the bottom
6. Saves the finished report to a local folder
7. Optionally launches the interactive HTML logistics map generator (`map_generator_.py`)

**Stack:** `win32com` · `tkinter` · `SAP GUI Scripting API` · `openpyxl` · `pandas` · `datetime` · `glob`

**Key design decisions:**
- Credentials are read from a local plaintext config file — never hardcoded
- Freight values are cached per contract base number (e.g. `MC-001-A` and `MC-001-B` share one freight entry)
- FX rate lookup uses the closest available prior date, not an exact match
- All company-specific T-codes, plant codes, and file paths are anonymized

---

### `map_generator_.py` — Interactive HTML Logistics Map Generator

**Problem solved:** After the sales report is generated, the logistics team needed a visual overview of active shipment locations — spreadsheets are not suitable for geographic analysis.

**What it does:**
- Reads the processed sales Excel report
- Extracts active contract locations, partner names, and commodity types
- Generates a self-contained interactive HTML map using Folium/Leaflet.js
- Color-codes markers by commodity and IncoTerm type
- Exports a single `.html` file that can be opened in any browser without dependencies

**Stack:** `pandas` · `folium` · `openpyxl`

---

### `Historical_Trade_Data_Analyzer.py` — Historical Trade Data Analyzer

**Problem solved:** Management needed period-over-period commodity price comparisons across multiple years of trade history stored in Excel files.

**What it does:**
- Loads multi-year trade data from Excel
- Calculates weighted average prices per configurable time period
- Flags deviations from historical baselines
- Exports a structured summary Excel file ready for presentation

**Stack:** `pandas` · `openpyxl`

---

## How to Run

```bash
pip install -r requirements.txt
python sap_sales_automation_.py
```

> Requires SAP GUI installed locally with Scripting API enabled.
> Credentials must be set in `~/sap_config.txt` on first run — the script creates a template automatically.

---

## Output Examples

> Sanitized screenshots will be added here. The output is a formatted `.xlsx` management report + optional `.html` logistics map.
