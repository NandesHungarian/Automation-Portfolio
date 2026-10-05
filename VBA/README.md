# VBA Automation Scripts

Excel and Outlook VBA macros built for **daily operations in agricultural commodity trading**. Each macro eliminates a specific manual, error-prone task from the daily workflow.

All scripts are anonymized — partner names, sheet names, and internal references have been replaced with generic placeholders.

---

## Scripts

### `Aviso_Automation.vba` — Daily Logistics Dispatch Email Generator

**Problem solved:** Every day the operations team had to manually copy shipment data from Excel into Outlook emails and send them to 5–15 logistics partners. This took ~1 hour per day and was prone to copy-paste errors.

**What it does:**
- Reads active shipment rows from the master Excel workbook
- Formats the data into an HTML table matching the company's standard aviso template
- Injects the table into an Outlook `.oft` template
- Sends individualized emails to each logistics partner automatically

**Time saved:** ~1 hour/day (~20 hours/month)

---

### `FCA_Price_Indication_Emailer.vba` — FCA Price Offer Distributor

**Problem solved:** After a morning price meeting, traders needed to send FCA price indications to counterparties — each covering a specific geographic region. The correct prices, products, and recipients had to be matched manually from a master table.

**What it does:**
- Reads the current price indication table from a structured Excel sheet
- Matches each row to the responsible regional coverage list
- Generates and sends personalised Outlook emails per counterparty with only their relevant price rows

---

### `Email_Attachment_Importer.vba` — Daily Report Import from Outlook

**Problem solved:** A daily data file arrived as an email attachment and had to be found in the inbox, saved, opened and copied into the working workbook by hand every morning.

**What it does:**
- Scans the most recent Outlook inbox items for today's email (subject prefix + `yyyymmdd` date tag)
- Saves the `.xlsx` attachment to a unique temp file and opens it read-only
- Replaces the contents of the target worksheet with the fresh data
- Reports success or a clear "not found" message; errors are caught instead of failing silently

---

### `Period_Average_Calculator.vba` — Weighted Average Price Calculator

**Problem solved:** Management reporting required weighted average commodity prices over configurable periods (weekly, monthly, crop year). This was done manually using complex nested Excel formulas that broke when source data structure changed.

**What it does:**
- Accepts a configurable date range input
- Calculates weighted averages (by quantity) for each commodity and base location combination
- Writes results to a formatted summary sheet

---

### `Tiered_Pricing_and_Deviation_Tracker.vba` — Price Deviation Auditor

**Problem solved:** Contracts were sometimes booked at prices outside the approved tiered pricing grid. These deviations were not caught until end-of-month reconciliation.

**What it does:**
- Reads the active tiered pricing grid
- Compares each contract's booked price against the applicable tier
- Highlights deviations with conditional formatting (color-coded by severity)
- Generates a summary audit trail sheet listing all flagged contracts

---

### `Incoming_Data_Consolidator.vba` — Multi-Sheet Data Merger

**Problem solved:** Multiple regional teams sent their data in separate Excel files or sheets with slightly different column orders. Consolidating them required manual copy-paste and deduplication.

**What it does:**
- Iterates over all source sheets (or files via folder path)
- Normalizes column order based on header names (not fixed column positions)
- Deduplicates by a unique contract identifier column
- Writes a clean master table to a target sheet

---

### `Regional_Coverage_Report_Distributor.vba` — Report Splitter & Distributor

**Problem solved:** The weekly master coverage report contained data for all regional coordinators. Manually splitting it and emailing each coordinator their slice took ~45 minutes every week.

**What it does:**
- Reads the master coverage sheet
- Groups rows by regional coordinator
- Creates a separate formatted sheet per coordinator
- Sends each sheet as an Outlook email attachment to the responsible person

**Time saved:** ~45 minutes/week (~3 hours/month)

---

## How to Use

Each `.vba` file contains the full macro code. To use in Excel:

1. Open the target Excel workbook
2. Press `Alt + F11` to open the VBA editor
3. Insert a new Module (`Insert → Module`)
4. Paste the script content
5. Adjust the sheet name and column references at the top of each script to match your workbook
6. Run with `F5` or assign to a button
