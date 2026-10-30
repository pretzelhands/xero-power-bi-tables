# xero-power-bi-tables

Turn Xero Profit and Loss exports into one long table that Power BI can load as is.

A Xero P&L export is laid out for reading. It has a title block, section headings, one column per month and total rows. Power BI wants one row per account and month instead. This script does that reshaping for you:

| Organisation | Section | Account | Period | Period start | Amount |
|---|---|---|---|---|---|
| Demo Company (Global) | Trading Income | Sales | Oct 2026 | 2026-10-01 | 8237.6 |
| Demo Company (Global) | Trading Income | Sales | Sep 2026 | 2026-09-01 | 9106.89 |
| Demo Company (Global) | Operating Expenses | Advertising | Oct 2026 | 2026-10-01 | 2309.47 |

Amounts keep the sign Xero shows, so income and expenses are both positive. Work out profit per section in Power BI (income minus cost of sales minus expenses) rather than summing the Amount column.

It leaves out headings, "Total ..." rows, Gross Profit, Net Profit and the Year to date column. Those are sums of other rows or columns, so keeping them would count the same money twice in a Power BI measure.

## Usage

You need Python 3.8 or newer. There is nothing to install.

1. In Xero, open the Profit and Loss and pick a range with one column per month, for example "Current financial year by month". Export it as Excel.
2. Run:

```
python3 xero_pl_to_table.py "Profit and Loss.xlsx" -o pl.csv
```

3. In Power BI Desktop, use Get data > Text/CSV and pick `pl.csv`. Set Period start to the Date type and Amount to Decimal number.

The script reports what it found on each file:

```
Demo Company (Global): 18 accounts x 4 columns. Left out Year to date (totals of the other columns)
Wrote 45 rows to pl.csv
```

### More than one company

Pass one export per company. The Organisation column comes from the company name at the top of each export, so you can filter or compare companies in one report:

```
python3 xero_pl_to_table.py north.xlsx south.xlsx -o group-pl.csv
```

If two files hold the same company and month (say a single-month export and a by-month export), the first file wins and the script tells you which months it skipped.

### Options

- `-o FILE` writes a CSV (UTF-8 with a byte order mark, so Excel opens it cleanly). Without it the table prints to the terminal.
- `--keep-zero` keeps rows where the amount is 0. By default they're dropped.

### What it reads

- `.xlsx` exports from Xero (the first sheet), and `.csv` files saved from one. The Excel export is the safer choice, because it marks Xero's own total rows. In a CSV the script can only go by row names. Comma, semicolon and tab separators all work, and so do amounts such as `1,234.50`, `(200.00)` and `1.234,50`.
- Columns split by tracking category (East, West) or compare periods work too. Period start stays blank for columns that aren't a month.
- Old `.xls` files don't work. Export again as Excel.

The `examples` folder has a real export of Xero's Demo Company and the table the script made from it.

## Other ways to do this

- If you'd rather not run a script, the same conversion runs in your browser at [reportpost.app/tools/xero-profit-and-loss-power-bi-table](https://reportpost.app/tools/xero-profit-and-loss-power-bi-table). Files stay on your computer.
- The [Xero to Power BI guide](https://reportpost.app/guides/xero-power-bi) compares the other routes: connector apps, the Xero API, and report files that Power BI refreshes from SharePoint or OneDrive.

## About

Made by [Reportpost](https://reportpost.app), which emails Xero reports to your team or clients on a schedule. Reportpost isn't affiliated with, made by or endorsed by Xero. "Xero" is used here only to say which exports the script reads.

MIT licence. Issues and pull requests are welcome.
