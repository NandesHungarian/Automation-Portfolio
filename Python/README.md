# Python Scripts

## Historical Trade Data Analyzer

`Historical_Trade_Data_Analyzer.py`

**The problem:** When sales had to be analysed over several years, every past contract had to be converted from HUF to EUR or USD with the forward rates that were valid on that day. These rates are stored in hundreds of daily Excel files, one file per day. Looking them up by hand for thousands of contracts was not realistic.

**The solution:** A small desktop app. You choose the SAP export with the contracts, and the script opens the daily rate files, matches every contract to the right day's rates and converts the prices. pandas keeps this fast even with several years of data.

**Key details:**
- Each contract is matched to the rate file of its document date, and within that file to the forward rate for one month after the shipment start
- The file names write the date in several ways (for example with or without leading zeros), and the script recognises all of them
- Flat price, truck freight and loading costs are all converted, so the result is a net price in EUR and USD
- Rows that would distort the averages are left out and marked with the reason, for example internal partners or delivered contracts without freight
- The work runs in the background, so the window stays responsive and shows the progress
- The result is an Excel file with every contract line and a summary sheet with weighted average prices per product

**Setup:** `pip install -r requirements.txt`, then set the folders in the configuration block at the top of the script.

---

## SAP Sales Report and Logistics Map

The daily SAP reporting script and the map generator have their own repository with full documentation:
[SAP_AgriTrade_Automation](https://github.com/NandesHungarian/SAP_AgriTrade_Automation)
