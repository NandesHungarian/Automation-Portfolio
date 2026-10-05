# Automation Portfolio

I worked in agricultural commodity trading, where a lot of the day went into opening emails and files, copying data between them and sending the same kind of email again and again. These are the tools I built to take that work over. They are written in Python and Excel VBA and connect SAP, Excel and Outlook.

**Together they save about 45 hours a month.**

Nándor Magyar · Malta · magyarnana97@gmail.com

---

## SAP sales report and logistics map

Python · [Repository and documentation](https://github.com/NandesHungarian/SAP_AgriTrade_Automation)

Every morning the sales report had to be pulled from SAP, converted to EUR with the day's exchange rates and formatted for management. This took 45 to 60 minutes. Now a Python script does it in a single run. It logs into SAP, exports the report, converts the prices, asks for any missing freight cost in a small window and builds the summary table. At the end it can draw an interactive map showing where the goods are going.

![Interactive logistics map](https://raw.githubusercontent.com/NandesHungarian/SAP_AgriTrade_Automation/main/images/logistics_map.jpg)

---

## Historical Trade Data Analyzer

Python · [Details](./Python)

For analysing sales over several years. To convert past sales from HUF to EUR or USD accurately, every contract needs the forward rates that were valid on its day. These are stored in hundreds of daily Excel files. The script opens them, matches every contract to the right day's rates and converts the prices. pandas keeps this fast even with years of data.

---

## Everyday Excel and Outlook macros

VBA · [Details for every macro](./VBA)

| Macro | What it does | Time saved |
|---|---|---|
| Aviso Automation | Prepares a daily email for each partner showing what they collected yesterday and what is still waiting | ~1 hour/day |
| Coverage Round Trip | Sends each regional colleague their part of the monthly coverage table, then brings the returned files back into the master table | 3–4 hours/month |
| Partner Group Prices | Sends daily prices to three partner groups, each with its own price tables, and logs them against the reference price | ~5 min/day |
| FCA Price Email | Sends the daily price table and saves the prices to a log sheet for later charts | ~5 min/day |
| Email Attachment Importer | Finds today's report email in Outlook and pastes its Excel attachment into the right sheet | ~5 min/day |
| Soybean Meal Period Averages | Connects two files that store periods differently, for the soy ratio | ~5 min each run |

---

## How I build these tools

- Emails are prepared but opened as drafts, so a person checks them before sending
- Data is found by labels, column headers or named cells, so the tools keep working when a table changes
- Settings and inputs stay in Excel, so colleagues can change them without touching the code
- They run on any colleague's computer, not only on mine
- When something is missing, the tool says exactly what, instead of skipping it quietly

---

All company data (partner names, SAP codes, folders, email addresses) has been anonymized. Screenshots were made with test data.
