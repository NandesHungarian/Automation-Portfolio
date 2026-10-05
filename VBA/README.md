# VBA Automation Scripts

Excel and Outlook macros I built for daily work in **agricultural commodity trading**. Each one replaces a specific manual task: opening emails and files, copying data between them, and preparing the same kind of email again and again.

All scripts are anonymized. Partner names, sheet names, folders and email addresses have been replaced with generic placeholders.

---

## Overview

| Macro | What it solves | Time saved |
|---|---|---|
| [Email Attachment Importer](#email_attachment_importervba--daily-report-import-from-outlook) | Finds today's report email and pastes its Excel attachment into the workbook | — |
| [Soybean Meal Period Averages](#period_average_calculatorvba--soybean-meal-averages-for-the-soy-ratio) | Connects two files that store periods differently, for the soy ratio | — |
| [FCA Price Email](#fca_price_indication_emailervba--daily-price-email-with-price-log) | Sends the daily price table and keeps a price log | — |
| [Tiered Pricing](#tiered_pricing_and_deviation_trackervba--daily-prices-to-three-partner-groups) | Sends daily prices to three partner groups and logs them against the reference price | — |
| [Aviso Automation](#aviso_automationvba--daily-pickup-summary-to-partners) | Tells each partner what they collected yesterday and what is still waiting | ~1 hour/day |
| [Coverage Round Trip](#monthly-coverage-round-trip) | Sends the monthly coverage table out to colleagues and brings their answers back | 3–4 hours/month |

---

## How I approach these problems

The same ideas come back in most of the macros.

- **A person checks before anything is sent.** The macros prepare the emails but open them as drafts. Someone looks at each one and clicks Send. *(Aviso, FCA Price Email, Tiered Pricing, Coverage Round Trip)*
- **It keeps working when the table changes.** Data is found by labels, column headers or Named Ranges instead of fixed cell addresses, so adding rows or moving data around does not break it. *(FCA Price Email, Soybean Meal Period Averages, Tiered Pricing)*
- **Inputs live in Excel, not in the code.** Prices, recipient lists and periods are named cells in the workbook. A colleague can update them without opening the VBA editor. *(Tiered Pricing, Soybean Meal Period Averages)*
- **It works on any colleague's machine.** Folder paths are built from the current Windows user name, and templates are searched in several folders. *(Aviso, Tiered Pricing)*
- **Nothing fails silently.** If something is missing, the macro says exactly what, instead of skipping it quietly. *(Aviso, Email Attachment Importer, Soybean Meal Period Averages, Coverage Round Trip)*
- **The data is kept for later.** Every price email also writes the day's prices to a log sheet, so price history is ready for analysis. *(FCA Price Email, Tiered Pricing)*

---

## Getting data in

### `Email_Attachment_Importer.vba` — Daily Report Import from Outlook

**The problem:** A daily report arrived as an Excel attachment. Every day someone had to find the email, save the file, open it and copy the data into the working workbook.

**The solution:** A button in the Excel workbook starts the macro. It opens classic Outlook, finds today's email by the date tag in its subject, opens the Excel attachment and pastes the data into the right worksheet.

**Key details:**
- Checks the latest 150 inbox emails for the subject prefix plus today's date (`yyyymmdd`)
- Skips meeting requests and other non-mail items
- Opens the attachment read-only and replaces the target sheet's contents with the fresh data
- Shows a clear message if today's email or its attachment is missing
- Easy to reuse. The subject and the target sheet are set in two lines at the top. I ran several copies of this macro for different daily emails, each with its own subject

---

### `Period_Average_Calculator.vba` — Soybean Meal Averages for the Soy Ratio

**The problem:** We compared sunflower meal and rapeseed meal prices to the soybean meal price (soy ratio). The soybean meal averages had to come from the daily soy calculation file for exactly the same periods as the sunflower and rapeseed prices, for example January to May, May to August or September to December. One file wrote the periods as text (e.g. "2026. január"), while the other stored the prices by date.

**The solution:** The macro connects the two files. It reads the periods as they are written, turns them into date ranges and calculates the soybean meal average for each. When a period changes, the averages follow.

**Key details:**
- Opens today's soy calculation file, or uses it if it is already open
- Sets the location in that file and recalculates it
- Understands Hungarian month names and covers whole months, so "január - május" runs until 31 May
- Loads the price data into memory in one step, so the calculation is fast
- Writes "no data" or "format error" next to a period instead of leaving a wrong number
- Puts the source file's inputs back to their original values at the end
- No fixed cell addresses. Single cells are Named Ranges and the date and price columns are found by their header

---

## Sending prices and notices

### `FCA_Price_Indication_Emailer.vba` — Daily Price Email with Price Log

**The problem:** The daily FCA prices had to be sent out every day. We also needed to look back later at how the prices had moved.

**The solution:** With one click the macro takes the price table as a picture and creates an Outlook email with the set recipient, subject and text, with the table in the body. It then saves the day's prices to a separate log sheet. This sheet works as a digital price diary. When we needed to show how prices had moved, the data was already there and could be turned into a chart.

**Key details:**
- The table goes into the email as an image, so it looks the same on every device
- Recipient, dated subject and standard text are filled in automatically
- The email opens as a draft for a final check
- Logs date, time, user and the prices of every product position, one row per day
- Finds each price row by a label in a helper column, not a fixed row number, so the log keeps working when rows are added or moved

---

### `Tiered_Pricing_and_Deviation_Tracker.vba` — Daily Prices to Three Partner Groups

**The problem:** Every day the factory prices went out to three partner groups (standard partners, brokers and end users), each with its own prices.

**The solution:** The macro creates one email per group from an Outlook template and pastes that group's two price tables, one per product, into the marked spots of the email. The recipients come from cells in the workbook and go into BCC, so partners do not see each other. After that it logs the day's prices for each delivery period together with the reference price and the difference, so we could follow how our premium moved over time.

**Key details:**
- Every input (price tables, recipient lists, period prices) is a named cell in the workbook. Updating prices or adding a new partner happens in Excel, never in the code
- Pastes the prices as real tables, not images, so partners can copy the numbers
- Retries the paste up to three times if the Windows clipboard is busy. This is a known cause of random errors (Error 4605) when copying from Excel into Outlook
- Finds the Outlook template in several possible folders
- The emails open as drafts for a final check
- Logs date, time, user, and price, reference price and difference for each delivery period (ASO, NDJ, FMA, MJJ)

---

### `Aviso_Automation.vba` — Daily Pickup Summary to Partners

**The problem:** The workbook tracked which partners collected their ordered goods and which did not. Every day each partner with an order for the previous day, usually 5–15 partners, had to get their own summary.

**The solution:** The macro goes through the list of these partners. For each one it filters the two pivot tables (collected and not collected) to that partner, opens the partner's own Outlook template (`.oft`) and places both tables into the marked spots in the email.

**Key details:**
- Finds each partner's template by name, searching several folders in order
- Converts the filtered tables into formatted HTML tables in the email
- Folder paths are built from the current Windows user name, so it runs on any colleague's machine
- The emails open as drafts. Someone looks over each one and clicks Send, so a wrong figure never reaches a partner unnoticed
- Partners without a template are skipped and listed at the end, so nobody is missed silently

**Time saved:** ~1 hour/day (~20 hours/month)

---

## Monthly coverage round trip

`Regional_Coverage_Report_Distributor.vba` + `Incoming_Data_Consolidator.vba`

**The problem:** A large master table showed each partner's coverage, meaning how much of their monthly need they had already bought. Every month each regional colleague got their own part of it to update with new figures. This was done by hand. More than 13 Excel files to open, a lot of copy-paste, and the same again when the answers came back.

**The solution:** Two macros that work as a pair. The first sends the files out, the second brings the answers back into the master table. This was my first macro.

**Part 1: `Regional_Coverage_Report_Distributor.vba` — Send out**
- Loops through the list of regional colleagues
- Opens each colleague's own file, filters the master table to that colleague's rows and copies them in
- Hides the helper columns so the file is clean to fill in
- Saves the file into a folder for the current month, created automatically if missing
- Prepares an Outlook email with the file attached and a reply deadline. At the end only the Send button has to be pressed

**Part 2: `Incoming_Data_Consolidator.vba` — Bring back**
- Goes through every returned file in the incoming folder
- Checks that the expected input sheet is there and reports any file that does not match
- Removes filters and unhides rows and columns, so nothing a colleague hid is lost
- Copies the raw data from all files into the master sheet as values
- Stops with a clear message if a file causes an error, instead of leaving half-imported data behind

**Time saved:** 3–4 hours every month, and no more copy-paste errors between 13+ files

---

## How to Use

Each `.vba` file contains the full macro code. To use it in Excel:

1. Open the target Excel workbook
2. Press `Alt + F11` to open the VBA editor
3. Insert a new module (`Insert → Module`) and paste the script
4. Set the values in the configuration block at the top of the script (folders, sheet names, Named Ranges, column headers)
5. Run with `F5` or assign the macro to a button

The macros that use Outlook need **classic Outlook for Windows**. The new Outlook does not support VBA.
