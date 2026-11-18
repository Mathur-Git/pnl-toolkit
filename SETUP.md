# SETUP: for Claude, setting this toolkit up on a new PC

You were asked to fetch this toolkit into a `pnl` folder and follow this file. The person trades futures through Wedbush; their daily and monthly statements arrive as PDF attachments by email, and the `pnl` folder held their Outlook export (`.pst`) of the mail folder they arrive in. Read `README.md` in full before doing anything else: it explains every part.

## 1. Check the tools

Git, Python 3, `pdfplumber` and `openpyxl` (`winget install Git.Git`, `pip install pdfplumber openpyxl` if missing). Nothing else is needed. If `python` opens the Microsoft Store, run the scripts with `py` instead; the double-click files in `start\` already prefer `py`.

## 2. Which case this is

- **`PnL Record.xlsx` is not in the folder** (a new setup): do 3.
- **`PnL Record.xlsx` is already there** (set up earlier from a copy of this toolkit): do 4.

## 3. A new setup

1. Run `python tools\outlook_import.py`. It moves the `.pst` into `Outlook-export\` and saves every statement PDF into `Daily-statements\` and `Monthly-statements\`. Tell them how many emails it read and how many daily and monthly statements it saved. If Outlook (classic) is open, it may say the file is locked: ask them to quit Outlook.
2. Run `python tools\update_record.py`, in the background: the first run reads every statement once into the ledger `data\statements.db` (about a second each), works out which account is which, makes `PnL Record.xlsx` from the template and writes the journal. If `PnL Record.xlsx` is open in Excel, ask them to close it.
3. Report back, briefly:
   - how many statements were read, and whether every one passed its checks. If any failed, show the reason the report gives, open that PDF, and work out what in the statement the reader did not expect before changing anything;
   - which accounts it found, and which it took as rates (SR3, SR1, ZQ) and as VIX. A product it does not know is flagged as unknown: add its statement description to `PRODUCTS` in `tools\ledger.py` and run again;
   - what the monthly check says. Each month should agree, except where daily statements are missing;
   - the latest day's account value, daily P&L, TC and contracts, from the report.
4. Ask them to open `PnL Record.xlsx` and check the `Statement` sheet (summary, monthly and quarterly tables, charts) against what they know of their year.

## 4. A folder set up earlier from a copy

1. Ask them to close `PnL Record.xlsx` if it is open, then copy it into `Backups\` with today's date in the name.
2. `git status` must show nothing: their statements, `data\` and the record are ignored by `.gitignore`. Anything listed is a question for them, never something to delete.
3. Rebuild the Statement sheet once in the new layout: `python tools\patch_statement.py`. Then run `python tools\update_record.py --dry` and show its report.
4. Tell them in two lines what is new (README, "The record").

## 5. Finish

Tell them in two lines their daily routine (README, "Every day"), and that double-clicking `start\0 Get latest.cmd` fetches every later change to the toolkit.

## Rules

- Never edit `PnL Record.xlsx` with openpyxl or anything that rewrites the workbook: the toolkit edits it through `tools\xlsxpatch.py` only (README, "For Claude").
- Never delete or rename a statement PDF or the `.pst`.
- The account summary at the top of the `Statement` sheet is laid out for a daily note to a manager. Leave it unless they ask for a different layout.

## Changing the toolkit

Everyone who uses the toolkit may improve it; every change reaches everyone through `start\0 Get latest.cmd`. Only when the person asks for a change:

1. `git pull --ff-only`, then a branch: `git switch -c <name>/<what>`. Never edit on `main`: a local change there makes `0 Get latest` refuse to update.
2. Make the change and check it on this folder's own record (`python tools\update_record.py --dry` first; a change to the Statement sheet: `python tools\patch_statement.py`, then open the record in Excel). After changing the Statement's layout or style, run `python tools\make_template.py` so new setups get it.
3. The repository is **public**: nothing committed may hold a figure from a statement, an account number, a name or a file path from this PC. Check `git diff --cached` for them before every commit; `.gitignore` already keeps the statements, `data\`, `Backups\` and the record out.
4. Push the branch and open a pull request on GitHub saying what changed and why. The person needs write access, given by the repository's owner (Settings > Collaborators); Git asks them to sign in to GitHub the first time.
5. Merge it once it is checked, then `git switch main` and `git pull --ff-only`. Never force-push and never rewrite `main`.
