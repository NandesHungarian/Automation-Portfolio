# VBA Automation Scripts

Excel and Outlook VBA macros built for daily operational use in agricultural commodity trading and logistics. Each script replaces a specific manual workflow that was either error-prone, time-consuming, or both.

All files use the `.vba` extension for readability in version control. To use them, open the VBA editor in Excel (`Alt+F11`) and paste the code into the appropriate module.

---

## Scripts

### [`Aviso_Automation`](./Aviso_Automation)

**What it does:** Generates and sends the daily logistics dispatch (aviso) emails to delivery partners without any manual input.

**The problem it solved:** Every morning, the logistics team had to open the dispatch plan, filter by partner, copy the relevant rows, format an email table manually, and send individual emails to each counterparty. One missed partner or copy-paste error meant a delayed truck.

**How it works:**
- Reads the day’s dispatch data from a structured Excel table
- Groups rows by delivery partner
- Builds an HTML-formatted table for each partner’s deliveries
- Injects the table into a pre-designed Outlook `.oft` template
- Sends each email automatically

**Time saved:** ~1 hour/day  
**Tech:** Excel VBA · Outlook VBA · HTML email templating

---

### [`FCA_Price_Indication_Emailer.vba`](./FCA_Price_Indication_Emailer.vba)

**What it does:** Distributes personalised FCA price indication emails to counterparties based on their regional coverage.

**The problem it solved:** Price indications had to go out to a list of partners each day, but each partner only receives prices for their specific coverage region. Manually filtering and sending individual emails was slow and created version-control risk (wrong prices to wrong partner).

**How it works:**
- Reads the price indication table and the partner–region mapping from Excel
- For each partner, filters the rows relevant to their coverage
- Composes and sends a personalised Outlook email with their specific price table

**Tech:** Excel VBA · Outlook VBA

---

### [`Period_Average_Calculator.vba`](./Period_Average_Calculator.vba)

**What it does:** Calculates weighted average commodity prices across configurable time windows.

**The problem it solved:** Management reports required period average prices (daily, weekly, monthly, crop-year) that had to be manually calculated from transaction data each time.

**How it works:**
- Reads raw transaction data (quantity, price, date)
- Calculates weighted averages for user-selected periods
- Writes results to a summary table formatted for management reporting
- Supports multiple commodities and delivery locations in one run

**Tech:** Excel VBA

---

### [`Tiered_Pricing_and_Deviation_Tracker.vba`](./Tiered_Pricing_and_Deviation_Tracker.vba)

**What it does:** Flags contracts where the agreed price deviates from tiered pricing thresholds, and generates an audit trail.

**The problem it solved:** With volume-based tiered pricing, identifying which contracts were priced outside their applicable threshold required manual cross-referencing of the contract list against the pricing matrix.

**How it works:**
- Reads the contract list and the tiered pricing matrix
- For each contract, identifies the applicable tier based on volume and partner category
- Compares the contracted price to the tier threshold
- Applies conditional formatting (green / amber / red) based on deviation magnitude
- Builds a summary audit sheet listing all out-of-band contracts

**Tech:** Excel VBA

---

### [`Incoming_Data_Consolidator.vba`](./Incoming_Data_Consolidator.vba)

**What it does:** Merges incoming data from multiple Excel sheets into a single normalised master table.

**The problem it solved:** Data from different partners arrived in slightly different Excel formats — different column orders, different header names, missing fields. Manually standardising and appending each file was tedious and introduced inconsistencies.

**How it works:**
- Reads all source sheets (or files) defined in a configuration table
- Maps each source column to the target master schema by name matching
- Deduplicates rows based on a configurable key column
- Appends all data to the master table with source tracking

**Tech:** Excel VBA

---

### [`Regional_Coverage_Report_Distributor.vba`](./Regional_Coverage_Report_Distributor.vba)

**What it does:** Splits a master coverage report by region and sends each section to the responsible regional coordinator via Outlook.

**The problem it solved:** The monthly coverage report had to be manually filtered per region, saved as a separate file, and emailed to each coordinator. With multiple regions, this was 20–30 minutes of repetitive work every month.

**How it works:**
- Reads the master report and the region–coordinator mapping table
- For each region, filters the relevant rows
- Writes the filtered data to a temporary sheet
- Attaches it to an Outlook email addressed to the correct coordinator and sends

**Tech:** Excel VBA · Outlook VBA

---

## Notes on File Extensions

The `Aviso_Automation` and `Email` files have no extension — they are plain VBA source files. Files ending in `.vba` are identically structured. All can be pasted directly into the Excel/Outlook VBA editor.
