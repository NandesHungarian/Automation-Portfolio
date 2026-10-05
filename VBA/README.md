# VBA Automation Scripts

Excel and Outlook VBA macros built for **daily operations in agricultural commodity trading**. Each macro eliminates a specific manual, error-prone task from the daily workflow.

All scripts are anonymized — partner names, sheet names, and internal references have been replaced with generic placeholders.

---

## Scripts

### `Aviso_Automation.vba` — Daily Pickup Summary Emails to Partners

**How it was used:** The workbook tracked which partners collected their ordered goods and which did not. The macro went through the list of partners who had orders for the previous day, usually 5–15 partners. For each one it filtered the two pivot tables (collected and not collected) to that partner, opened the partner's own Outlook template (`.oft`) and placed both tables into the marked spots in the email (`{Tábla1}`, `{Tábla2}`). Partners without a template were skipped and listed at the end, so nobody was missed silently.

**Built-in check:** The emails open as drafts instead of being sent straight away. Someone looks over each one and clicks Send, so a wrong figure never reaches a partner unnoticed.

**What it does:**
- Loops through the previous day's partner list and filters both pivot tables per partner
- Finds the partner's template, searching several folders in order
- Converts the filtered ranges into formatted HTML tables
- Works on any colleague's machine, because the template folders are built from the current Windows user name
- Shows a summary at the end with the drafts created and the partners that need attention

**Time saved:** ~1 hour/day (~20 hours/month)

---

### `FCA_Price_Indication_Emailer.vba` — Daily Price Email with Built-in Price History

**How it was used:** The daily FCA prices were kept in a table in Excel. With one click the macro took this table as a picture and created a new Outlook email with the set recipient, subject and text, with the price table inserted in the body. It then saved the day's prices to a separate log sheet. This sheet worked as a digital price diary. When we later needed to know how prices had moved, the answer was already there and could be turned into a chart in minutes.

**What it does:**
- Exports the price table range as an image and places it inline in the email body, so it looks the same on every device
- Fills in recipient, dated subject and the standard text automatically
- Opens the email as a draft for a final check before sending
- Logs date, time, sender and every product position's prices to the log sheet, one row per day
- Finds each price row by a label in a helper column, not a fixed row number, so the log keeps working when rows are added or moved in the price table

---

### `Email_Attachment_Importer.vba` — Daily Report Import from Outlook

**How it was used:** A button in the Excel workbook started the macro. It opened classic Outlook, looked for today's email by the date tag in its subject, opened the Excel attachment and pasted the data into the right worksheet. Before, this meant finding the email, saving the file, opening it and copying the data by hand every day.

**Easy to reuse:** The subject and the target sheet are set in two lines at the top. I ran several copies of this macro for different daily emails, each with its own subject.

**What it does:**
- Checks the latest 150 inbox emails for the subject prefix plus today's date (`yyyymmdd`)
- Skips meeting requests and other non-mail items
- Saves the `.xlsx` attachment to a temp file and opens it read-only
- Replaces the target sheet's contents with the fresh data
- Shows a clear message if today's email or its attachment is missing

**Requirement:** classic Outlook for Windows. The new Outlook does not support VBA.

---

### `Period_Average_Calculator.vba` — Soybean Meal Averages for the Soy Ratio

**How it was used:** We compared sunflower meal and rapeseed meal prices to the soybean meal price (soy ratio). For this, the soybean meal averages had to come from the daily soy calculation file for exactly the same periods as the sunflower and rapeseed prices, for example January to May, May to August or September to December. The difficulty was that one file wrote the periods as text (e.g. "2026. január") while the other stored the prices by date. The macro connects the two. When the period changes, the averages follow.

**What it does:**
- Opens today's soy calculation file, or uses it if it is already open
- Sets the location in that file and recalculates it
- Turns the text periods, with Hungarian month names, into real date ranges covering whole months
- Averages the daily prices in each period, loading the data into memory in one step so it runs fast
- Writes each average next to its period, with a clear note if data is missing or a period is written wrongly
- Puts the source file's inputs back to their original values at the end
- Has no hard-coded cells. Single cells are referenced through Named Ranges and the date and price columns are found by their header, so inserting rows or columns does not break it

Files that describe the same thing in different ways are common. This macro is one example of keeping them in sync without retyping anything.
---

### `Tiered_Pricing_and_Deviation_Tracker.vba` — Daily Prices to Three Partner Groups, with Price Log

**How it was used:** Every day the factory prices went out to three partner groups (standard partners, brokers and end users), each with its own prices. The macro created one email per group from an Outlook template and pasted that group's two price tables, one per product, into the marked spots of the email. The recipients came from cells in the workbook and went into BCC, so partners did not see each other. After that it logged the day's prices for each delivery period together with the reference price and the difference, so we could follow how our premium moved over time.

**Easy for anyone to use:** Every input (price tables, recipient lists, period prices) is a named cell in the workbook. Updating prices or adding a new partner happens in Excel, never in the code.

**What it does:**
- Finds the Outlook template in several possible folders, so it works on any colleague's machine
- Pastes the prices as real formatted tables, not images, so partners can copy the numbers
- Retries the paste up to three times if the Windows clipboard is busy, a common cause of random errors (Error 4605) when copying from Excel into Outlook
- Opens the emails as drafts for a final check before sending
- Logs date, time, user, and price, reference price and difference for each delivery period (ASO, NDJ, FMA, MJJ)
---

### Monthly Coverage Round Trip — `Regional_Coverage_Report_Distributor.vba` + `Incoming_Data_Consolidator.vba`

These two macros work as a pair. The first sends the coverage files out, the second brings the answers back.

**Background:** A large master table showed each partner's coverage, meaning how much of their monthly need they had already bought. Every month each regional colleague got their own part of it to update with new figures. Before, this was done by hand. More than 13 Excel files to open, a lot of copy-paste, and the same again when the answers came back. This was my first macro.

#### Part 1: `Regional_Coverage_Report_Distributor.vba` — Send out

- Loops through the list of regional colleagues
- Opens each colleague's own file, filters the master table to their partners and copies the rows in
- Hides the helper columns so the file is clean to fill in
- Saves the file into a folder for the current month, created automatically if missing
- Prepares an Outlook email with the file attached and a reply deadline. At the end only the Send button has to be pressed

#### Part 2: `Incoming_Data_Consolidator.vba` — Bring back

- Goes through every returned file in the incoming folder
- Checks that the expected input sheet is there and reports any file that does not match
- Removes filters and unhides rows and columns, so nothing a colleague hid is lost
- Appends the raw data from all files into the master sheet as values
- Stops with a clear message if a file causes an error, instead of leaving half-imported data behind

**Time saved:** 1–2 hours every month, and no more copy-paste errors between 13+ files
---

## How to Use

Each `.vba` file contains the full macro code. To use in Excel:

1. Open the target Excel workbook
2. Press `Alt + F11` to open the VBA editor
3. Insert a new Module (`Insert → Module`)
4. Paste the script content
5. Adjust the sheet name and column references at the top of each script to match your workbook
6. Run with `F5` or assign to a button
