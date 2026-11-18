# P&L record from Wedbush statements

Turns the daily Wedbush statement emails into two things: a **ledger** (`data\statements.db`) that keeps everything each statement says, including every fill, and a **P&L record** in Excel built from that ledger. Each PDF is read once, when it first arrives; nothing reads it again.

## What one update does

1. **Outlook export to PDFs** (`start\1 Import from Outlook.cmd` only): takes the newest Outlook export in `Outlook-export\` and saves any statement PDF not already in the folder. It never deletes or overwrites a PDF.
2. **New PDFs into the ledger**: each PDF the ledger hasn't seen (known by its contents' fingerprint, not its name) is read, checked against its own arithmetic, and stored. A statement that fails a check is not stored, and the report says why.
3. **Ledger into the journal**: every stored day is written onto its row of `PnL Record.xlsx` (sheet `IMR+PnL`), overwriting what is there: the statements are the source of truth. The `Statement` sheet updates itself from the journal. The record is copied to `Backups\` first, and nothing is saved if nothing changed.
4. **A short report**: a few lines saying what is new, what was written, and anything that needs a look. For example:

```
Statements: 1 new (17 Nov).
Journal 17 Nov: A/c Val 8,240.50, Daily PnL +1,215.30, VIX value -210.40, TC 96.20, 180 contracts.
Monthly check: 11 of 11 agree (nothing new).
Saved PnL Record.xlsx.
```

The full detail of every run (each cell changed, each monthly comparison) goes to `data\run_log.txt`, which you never need to open unless something looks wrong.

## Every day

1. In **classic Outlook** (Start menu: "Outlook (classic)"; the new Outlook cannot do this): File > Open & Export > Import/Export > Export to a file > Outlook Data File (.pst). Pick the folder the statement emails are in, untick "Include subfolders", save it into `Outlook-export\` (any name), no password. Quit classic Outlook afterwards: it locks the file while open.
2. Close `PnL Record.xlsx` if it is open.
3. Double-click **`start\1 Import from Outlook.cmd`**.

Saved a statement PDF yourself? Put it in `Daily-statements\` and double-click **`start\2 Update from PDFs.cmd`**. **`start\Check only (writes nothing).cmd`** shows what an update would do without writing. Re-running anything is always safe.

**`start\0 Get latest.cmd`** fetches the newest toolkit from GitHub. It changes only `start\`, `tools\` and this README, never your statements, ledger or record; the record takes any new look at its next update.

## The folder

| Item | What it is |
|---|---|
| `PnL Record.xlsx` | the record: `IMR+PnL` is the daily journal, `Statement` the summary built from it |
| `start\` | the double-click files |
| `Outlook-export\` | the Outlook `.pst` export |
| `Daily-statements\` | `statementYYYYMMDD.pdf`, one a day; `duplicates\` holds a second copy of a day if Wedbush sent two |
| `Monthly-statements\` | `mstatementYYYYMMDD.pdf`, one a month |
| `data\statements.db` | the ledger (below) |
| `data\run_log.txt` | the full detail of every run |
| `Backups\` | a copy of the record before every save; the newest 10 are kept |
| `tools\` | the Python scripts; `tools\template\` holds a blank record |

## The ledger: `data\statements.db`

One SQLite file, the same kind stir-desk uses for its own data. Open it with any SQLite browser, or from Python:

```python
import sqlite3, pandas as pd
con = sqlite3.connect(r"data\statements.db")
pd.read_sql("SELECT * FROM fills WHERE date = '2026-11-13'", con)
```

Dates are ISO (`2026-11-13`); amounts are as the statement states them, so costs and losses are negative.

| Table | One row per | Holds |
|---|---|---|
| `files` | PDF read | file, daily or monthly, date, fingerprint, ok or failed (and why), whether it was used (0 for a second copy of a day that was not) |
| `accounts` | account | which account is which: `rates` (trades SR3, SR1, ZQ) or `vix` (trades VX), worked out from what each trades |
| `summary` | day, account, line | every line of the account summary: beginning balance, commission, exchange fee, clearing fee, total commission and fees, gross profit or loss, ending balance, open trade equity, total equity, account value at market, initial and maintenance margin; and the month-to-date and year-to-date realized P&L, commissions and fees |
| `fill_blocks` | day, account, contract | contracts bought and sold, average buy and sell price, commission, exchange fee, clearing fee, other fees, total: what drove the day's TC |
| `fills` | fill | trade date, product, contract (e.g. `MAR 28`), side (B or S), quantity, price, whether it was a spread leg |
| `realized` | closing line | product, contract, quantity, gross P&L |
| `positions` | day, account, contract | open position at the close: side (L or S), quantity, average price, settle, open P&L |
| `adjustments` | adjustment | description (e.g. a CME fee refund), fee line, amount (positive = money back) |
| `monthly` | month, account, line | the monthly statement's lines, as in `summary` |

What the ledger is for: reconciling positions, seeing what was traded each day and what each structure (spread, fly, double play) cost, and later matching live TT fills against the broker's record.

## Checks

A statement is stored only if all of these hold:

- each fill group: its commission plus fees equals its own "TOTAL COMM AND FEE"; its fills' bought and sold quantities equal its own counts; and each side's average fill price equals the statement's own average buy or sell price. Buy and sell are read from the column a quantity sits in, so this confirms every fill's side;
- each account: the fill groups plus adjustments equal "TOTAL COMMISSION AND FEES"; the closing lines add up to "GROSS PROFIT OR LOSS"; the open positions add up to "OPEN TRADE EQUITY"; beginning balance plus fees plus gross P&L equals the ending balance; ending balance plus open trade equity equals account value at market (a figure read off the wrong row breaks this); the three currency columns agree;
- the date inside the PDF matches its file name.

A day with two copies: both are read; they must agree on account values, and the copy that states the margin is used (one 9 Jan 2026 copy leaves it out).

The monthly check compares each monthly statement's month-end account value and the month's TC with the daily statements, per account. The report shows it only when a monthly statement is new or something disagrees without a known reason.

## The record

The journal's columns (sheet `IMR+PnL`, one row a weekday, dated through 31 Dec of the latest statement's year):

| Column | Holds |
|---|---|
| B, C | day, date |
| D, E | rates account value at market (cumulative P&L), initial margin |
| F | Daily PnL: D minus the nearest earlier row with a D |
| G, H | VIX account value, margin (blank when the VIX account is not on the statement; margin 0 when it is but holds nothing) |
| I | VIX Daily PnL, as F over G |
| J:L | SR3 TC, SR1 TC, ZQ TC: the rates account's fills |
| M | VX TC: everything the VIX account paid |
| N | Rates TC Adj: the rates account's refunds or charges. Rates TC = J+K+L+N |
| O | Total TC, equal to the statements' own total |
| P:S | SR3 Vol, SR1 Vol, ZQ Vol, VX Vol: contracts bought plus sold, every spread leg counted |

If an account is missing from the daily statements for part of the year, its TC for that time is taken from the monthly statements and added to M (VIX) or N (rates) at that month end, with the account's first P&L.

The `Statement` sheet: **A/C summary** (rows 4 to 11), the block copied and sent every day, with nothing beside it so a copy picks up only that block, and the typed adjustment under it. Its P&L lines and account value are the rates and VIX accounts together; its margin is the rates account's. Then **Monthly** (Jan to Dec, then Year), **Quarterly** and **By year**, all with the same columns, total first: Total P&L, Rates P&L, VIX P&L | Total TC, Rates TC, VIX TC | Total Volume, Rates Volume, VIX Volume. Figures show to the unit; cells keep exact amounts. Total columns are at full size, Rates and VIX in a smaller font; no figure is ever red. Below the tables, three **charts**: Total P&L year to date (a line from 1 Jan to 31 Dec, darker vertical lines at each quarter's start), and Total P&L and Total TC by month side by side. The year chart reads the **CHART DATA** rows at the bottom (one row a calendar day, hidden in a group: click the + beside the bar to see them). Every update rebuilds this sheet and its charts, so formatting typed on it by hand, or a chart moved by hand, does not last: change `tools\patch_statement.py` (layout) or `tools\statement_charts.py` (the charts' look) instead. P&L is after TC. A VIX cell shows "-" when no statement in the period carried the VIX account. Analysis beyond sums belongs in Python, reading the ledger.

## Setting it up in another folder (e.g. a colleague's)

The toolkit lives at https://github.com/Mathur-Git/pnl-toolkit. Only `start\`, `tools\`, this README, `SETUP.md` and the git settings are in it; `.gitignore` keeps everything else (the `.pst`, statements, `data\`, `Backups\`, `PnL Record.xlsx`) on each person's own PC.

The person puts their Outlook `.pst` export in a `pnl` folder and asks their Claude to fetch the toolkit into it (`git init -b main`, `git remote add origin https://github.com/Mathur-Git/pnl-toolkit.git`, `git fetch origin`, `git reset --hard origin/main`, `git branch -u origin/main`: this adds the toolkit's files and leaves everything else in the folder alone) and follow `SETUP.md`. The first run makes the folders, reads every statement into a new ledger (about a second each), works out which account is which, and makes `PnL Record.xlsx` from the template, starting on the first statement's date. How anyone using the toolkit changes it (a branch and a pull request; the repository is public) is in `SETUP.md`, "Changing the toolkit".

Account numbers are never typed. A product other than SR3, SR1, ZQ or VX is flagged as unknown: add its statement description to `PRODUCTS` in `tools\ledger.py`.

## If something fails

- **"is open in Excel"**: close the workbook and run again.
- **"The .pst is locked"**: quit classic Outlook completely.
- **A statement could not be read**: the report gives the check that failed. Open that PDF and compare. A new kind of line or section is the usual cause; the patterns are at the top of `tools\ledger.py`. Once fixed, delete that file's row from the `files` table (or the whole `statements.db`, to rebuild) and run again.
- **Monthly check disagrees**: usually a day missing from the daily statements.

## For Claude

- `tools\ledger.py` reads PDFs into `data\statements.db`; `tools\update_record.py` writes the journal from the ledger and prints the report; `tools\monthly_check.py` reconciles against the monthly statements; `tools\outlook_import.py` pulls PDFs from a `.pst` (reader in `tools\wedbush-extract\`; never use that folder's `extract.py --export-pdfs` on `Daily-statements\`, it deletes the PDFs in its target first).
- Never edit `PnL Record.xlsx` with openpyxl: it drops charts and controls. Edits go through `tools\xlsxpatch.py`. `tools\patch_style.py` holds the house style and runs on every save. `tools\patch_statement.py` owns the whole Statement sheet, its drawing and charts included (found through the sheet's rels, since Excel renames chart parts when it saves), and rebuilds it on every save; after changing it, run `tools\make_template.py`. Never refer to a style by its number: Excel renumbers fonts, borders, fills and cell styles when someone saves the record, so `patch_style.house()` first puts the house ones back in front and keeps the rest, and cell styles are looked up by their XML.
- Removing formula cells: drop `xl/calcChain.xml` or Excel refuses the file (`patch_style.apply` does this on every save).
- Later home: the desk's `stir-desk` repo, with code under its `lib/`, launchers under its `start/`, and the ledger in its `data/databases/`, out of git.
