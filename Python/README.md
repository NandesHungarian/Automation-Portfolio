# Python Automation Scripts

This folder contains Python scripts for data processing, SAP ERP integration, and automated report generation. All scripts are built for Windows environments in a corporate agricultural trading context.

---

## Scripts

### `sap_sales_automation_.py`
**SAP Sales & Logistics Workflow Automation**

End-to-end automation pipeline that connects to a live SAP ERP system via COM-based GUI scripting, exports the daily or weekly sales report, and processes it through a multi-step transformation chain:

- **SAP Login & Export:** Attaches to a running SAP GUI session (or launches one), navigates to the custom transaction code, sets the date range, and exports the ALV report to Excel
- **Multi-currency EUR Conversion:** Reads daily price indication files from a shared folder, performs date-matched FX lookups (HUF/USD → EUR) and calculates `Flat EUR`, `freight EUR`, and `Net Flat EUR` per row
- **Missing Freight Handling:** If a contract row requires freight cost but none is recorded, a `tkinter` popup prompts the user in real-time — the entered value is cached and backfilled to all related rows automatically
- **Partner Filtering:** Excludes internal intercompany partners and flags suspicious rows (e.g. HUF price below threshold, DDP/CPT with no freight) with red row highlighting
- **Management Summary Table:** Draws a pivot-style summary by location × commodity × crop year at the bottom of the report, with Friday-only breakdown for weekly runs
- **Map Generation Prompt:** Offers to launch `map_generator_.py` after report completion

**Stack:** `win32com` · `tkinter` · `ttk` · `openpyxl` · `datetime` · `glob` · `shutil`  
**Platform:** Windows only (SAP GUI + COM automation)

> All T-codes, system names, user IDs, file paths, and partner names have been anonymized.

---

### `Historical_Trade_Data_Analyzer.py`
**Historical Trade Data Analyzer**

Analyses multi-year commodity trade history exported from ERP systems. Designed to work with structured Excel exports.

- Calculates weighted average prices across configurable period windows
- Detects and flags price deviations from historical baselines
- Exports clean, structured Excel summaries formatted for management reporting

**Stack:** `pandas` · `openpyxl`

---

## Requirements

See the root [`requirements.txt`](../SAP/) or individual script headers for dependencies.  
All scripts require Python 3.9+.

```
pip install pandas openpyxl pywin32
```

> `pywin32` (`win32com`) is Windows-only and required for both SAP GUI scripting and Excel COM automation.
