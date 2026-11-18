"""One update: read any new statement PDFs into the ledger (data/statements.db), write the journal
(IMR+PnL) of PnL Record.xlsx from the ledger, and print a short report. The statements are the
source of truth: every stored day is written onto its journal row, overwriting what is there.

    D  A/c Val        rates account value at market (cumulative P&L)
    E  Margin         rates account initial margin
    F  Daily PnL      =D(this row) - D(nearest earlier row with a D)
    G  VIX A/c Val    VIX account value at market (only when the VIX account is on the statement)
    H  VIX Margin     VIX initial margin (0 when on the statement with no margin line)
    I  VIX Daily PnL  =G(this row) - G(nearest earlier row with a G)
    J:L  SR3 TC, SR1 TC, ZQ TC: the rates account's commission + every fee on its fills, positive
    M    VX TC: everything the VIX account paid (fills and adjustments), plus VIX TC the daily
         statements lack, from the monthly statements (monthly_check.carry_ins)
    N    Rates TC Adj: the rates account's refunds (negative) or charges (plus any rates TC the
         daily statements lack). Rates TC = J+K+L+N, VIX TC = M
    O    Total TC = SUM(J:N), equal to the statements' TOTAL COMMISSION AND FEES
    P:S  SR3 Vol, SR1 Vol, ZQ Vol, VX Vol: contracts bought + sold, every spread leg counted

The report says what is new, what was written, and only what needs a look; the full detail of
every run is appended to data/run_log.txt. No PnL Record.xlsx yet: one is made from
tools/template/, its first date set to the first statement's. Edits go through tools/xlsxpatch.py
(never openpyxl); the record is copied to Backups/ before a save (newest 10 kept); nothing is
saved if nothing changed; it refuses to run while the workbook is open in Excel.

    python tools/update_record.py          write
    python tools/update_record.py --dry    show what would change, write nothing
"""
import os
import re
import sys
from datetime import date, datetime, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ledger  # noqa: E402
import monthly_check  # noqa: E402
import patch_statement  # noqa: E402
import patch_style  # noqa: E402
from paths import DATA, RECORD, TEMPLATE, backup, record_is_open  # noqa: E402
from xlsxpatch import Book  # noqa: E402

SHEET = "IMR+PnL"
FIRST_ROW = 3
NUM_STYLE = "3"
DAY_STYLE, DATE_STYLE = "1", "2"
NUMBERS = {"J": "cost_SR3", "K": "cost_SR1", "L": "cost_ZQ", "M": "cost_VX", "N": "cost_ADJ",
           "P": "vol_SR3", "Q": "vol_SR1", "R": "vol_ZQ", "S": "vol_VX"}
INTS = {"P", "Q", "R", "S"}
EPOCH = date(1899, 12, 30)
LOG = os.path.join(DATA, "run_log.txt")


def cached(raw):
    """(formula, value) of a cell's XML; either may be None."""
    if not raw:
        return None, None
    f = re.search(r"<f[^>]*>(.*?)</f>", raw, re.S)
    v = re.search(r"<v>(.*?)</v>", raw, re.S)
    return (f.group(1) if f else None), (v.group(1) if v else None)


def as_num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def value_of(raw):
    """A cell's cached number, or None for a blank or a text cell (e.g. a typed note)."""
    if not raw or re.search(r'\st="(?:s|str|inlineStr|e|b)"', raw):
        return None
    return as_num(cached(raw)[1])


def fmt(x):
    x = round(float(x), 2)
    return str(int(x)) if x == int(x) else repr(x)


def journal_dates(sh):
    """date -> row. Typed dates are read as they are; a row with the roll-forward formula
    (=C(n-1)+IF(B(n-1)="Fri",3,1)) is the previous date + 1 day, + 3 after a Friday."""
    rows, d, r = {}, None, FIRST_ROW
    while sh.raw(f"C{r}"):
        f, v = cached(sh.raw(f"C{r}"))
        if f is None and as_num(v) is not None:
            d = EPOCH + timedelta(days=int(float(v)))
        elif f is not None and d is not None:
            d += timedelta(days=3 if d.weekday() == 4 else 1)
        else:
            break
        rows[d] = r
        r += 1
    return rows


def short(d):
    return date.fromisoformat(d).strftime("%d %b").lstrip("0")


def main(dry):
    if record_is_open() and not dry:  # a check-only run writes nothing, so it may run anyway
        print("PnL Record.xlsx is open in Excel. Close it and run again.")
        return 1
    log = [f"===== {datetime.now():%Y-%m-%d %H:%M:%S}{' (dry run)' if dry else ''}"]
    rep = ledger.update()
    days = ledger.daily_rows()

    # ---- report: statements
    new = rep["new_days"]
    line = f"Statements: {len(new)} new" + (f" ({', '.join(short(d) for d in new)})" if 0 < len(new) <= 5
                                           else f" ({short(new[0])} to {short(new[-1])})" if new else "")
    print(line + ("." if not rep["failed"] else f"; {len(rep['failed'])} could not be read:"))
    for f, why in rep["failed"]:
        print(f"  !! {f}: {why}")
    for day, used, dropped in rep["replaced"]:
        print(f"  {short(day)} came twice; using {used} (it states the margin), not {dropped}.")
    if rep["new_months"]:
        print(f"Monthly statements: {len(rep['new_months'])} new ({', '.join(m[:7] for m in rep['new_months'])}).")
    log += [line] + [f"failed {f}: {w}" for f, w in rep["failed"]]
    if not days:
        print("No statements in the ledger yet.")
        return 1

    # ---- TC by account: M = everything the VIX account paid; N = the rates account's adjustments
    for d in days:
        d["cost_VX"] = d["tc_vix"]
        d["cost_ADJ"] = round(d["tc_rates"] - d["cost_SR3"] - d["cost_SR1"] - d["cost_ZQ"], 2)
    by_date = {d["date"]: d for d in days}
    for c in monthly_check.carry_ins():
        if c["date"] in by_date:
            k = "cost_VX" if c["role"] == "vix" else "cost_ADJ"
            by_date[c["date"]][k] = round(by_date[c["date"]][k] + c["amount"], 2)
            log.append(f"carry-in {c['date']} +{c['amount']:,.2f}: {c['why']}")

    fresh = not os.path.exists(RECORD)
    book = Book(TEMPLATE if fresh else RECORD)
    sh = book.sheet(book.sheet_parts()[SHEET], SHEET)
    fails = sh.expand_shared()
    if fails:
        print("Could not expand shared formulas; nothing written:\n  " + "\n  ".join(fails))
        return 1
    if fresh:
        first = date.fromisoformat(days[0]["date"])
        sh.set_number(f"C{FIRST_ROW}", str((first - EPOCH).days), style=DATE_STYLE)
        print(f"New PnL Record.xlsx made from the template, journal starting {short(days[0]['date'])}.")

    row_of = journal_dates(sh)
    last, last_date = max(row_of.values()), max(row_of)
    changes = []

    def put_num(ref, value):
        f, v = cached(sh.raw(ref))
        if f is None and v is not None and abs(float(v) - float(value)) < 0.005:
            return
        changes.append((ref, f"={f}" if f else (v if v is not None else "(blank)"), fmt(value)))
        sh.set_number(ref, fmt(value), style=NUM_STYLE)

    def put_formula(ref, formula):
        f, v = cached(sh.raw(ref))
        if f == formula:
            return
        changes.append((ref, f"={f}" if f else (v if v is not None else "(blank)"), "=" + formula))
        sh.set_formula(ref, formula, style=NUM_STYLE)

    year_end = date(date.fromisoformat(days[-1]["date"]).year, 12, 31)  # dated through 31 Dec
    added_rows = 0
    while last_date + timedelta(days=3 if last_date.weekday() == 4 else 1) <= year_end:
        r = last + 1
        sh.set_formula(f"B{r}", f'TEXT(C{r},"DDD")', style=DAY_STYLE)
        sh.set_formula(f"C{r}", f'C{last}+IF($B{last}="Fri",3,1)', style=DATE_STYLE)
        last_date += timedelta(days=3 if last_date.weekday() == 4 else 1)
        row_of[last_date] = r
        sh.rows.setdefault(r, [f' r="{r}"', {}])
        last, added_rows = r, added_rows + 1

    d_val = {r: value_of(sh.raw(f"D{r}")) for r in range(FIRST_ROW, last + 1)}
    g_val = {r: value_of(sh.raw(f"G{r}")) for r in range(FIRST_ROW, last + 1)}
    written = {}
    for d in days:
        r = row_of.get(date.fromisoformat(d["date"]))
        if r is None:
            print(f"  !! {short(d['date'])}: no journal row has this date. Skipped.")
            continue
        written[r] = d["date"]
        put_num(f"D{r}", d["main_acct_value"])
        d_val[r] = d["main_acct_value"]
        put_num(f"E{r}", d["main_init_margin"])
        if d["vix_acct_value"] is not None:
            put_num(f"G{r}", d["vix_acct_value"])
            g_val[r] = d["vix_acct_value"]
            put_num(f"H{r}", d["vix_init_margin"] or 0)
        for col, key in NUMBERS.items():
            put_num(f"{col}{r}", int(d[key]) if col in INTS else d[key])
        put_formula(f"O{r}", f"SUM(J{r}:N{r})")
    prev = {}
    for r in written:
        for out, src, vals in (("F", "D", d_val), ("I", "G", g_val)):
            if vals.get(r) is None:
                continue
            p = next((q for q in range(r - 1, FIRST_ROW - 1, -1) if vals.get(q) is not None), None)
            if out == "F":
                prev[r] = p
            if p:
                put_formula(f"{out}{r}", f"{src}{r}-{src}{p}")
            elif out == "F":
                put_formula(f"{out}{r}", f"{src}{r}")

    # ---- report: journal
    new_rows = {r for r, d in written.items() if d in new}
    other = [c for c in changes if int(re.sub(r"\D", "", c[0])) not in new_rows]
    log += [f"{ref}: {old} -> {nv}" for ref, old, nv in changes]
    if len(new_rows) > 5:  # a first run or a long gap: one line for the lot, then the latest day
        print(f"Journal: {len(new_rows)} days written ({short(written[min(new_rows)])} to "
              f"{short(written[max(new_rows)])}). Latest:")
    for r in sorted(new_rows)[-1:] if len(new_rows) > 5 else sorted(new_rows):
        d = by_date[written[r]]
        pnl = d_val[r] - (d_val[prev[r]] if prev.get(r) else 0)
        tc = sum(d[k] for k in ("cost_SR3", "cost_SR1", "cost_ZQ", "cost_VX", "cost_ADJ"))
        n = sum(d[f"vol_{p}"] for p in ledger.PRODUCT_ORDER)
        vix = f", VIX value {d['vix_acct_value']:,.2f}" if d["vix_acct_value"] is not None else ""
        print(f"Journal {short(written[r])}: A/c Val {d['main_acct_value']:,.2f}, Daily PnL {pnl:+,.2f}"
              f"{vix}, TC {tc:,.2f}, {n:,} contracts.")
    if other:
        print(f"Journal: {len(other)} other cell(s) changed (details in data\\run_log.txt):")
        for ref, old, nv in other[:8]:
            print(f"  {ref}: {old} -> {nv}")
    if not changes and not added_rows and not fresh and not new_rows:
        print("Journal: already up to date.")

    # ---- report: monthly check, only what is new or unexplained
    res = monthly_check.check()
    shown = [x for x in res if not x["ok"] and (x["month_end"] in rep["new_months"] or not x["known"])]
    agree = sum(x["ok"] for x in res)
    known = sum(not x["ok"] and x["known"] for x in res)
    print(f"Monthly check: {agree} of {len(res)} agree" + (f", {known} a known gap already filled" if known else "")
          + (f"; {len(rep['new_months'])} new monthly statement(s)." if rep["new_months"] else " (nothing new)."))
    for x in shown:
        print(f"  {x['month_end'][:7]} {x['account']}: {x['result']}")
    log += [f"monthly {x['month_end']} {x['account']}: {x['result']}" for x in res]

    if dry:
        print("Dry run: nothing written.")
    elif changes or added_rows or fresh:
        patch_statement.apply(book)  # every save rebuilds the Statement, so a new layout reaches every record
        patch_style.apply(book)
        book.sub("xl/workbook.xml", r"<calcPr ((?:(?!fullCalcOnLoad)[^/])*?)/>",
                 r'<calcPr \1 fullCalcOnLoad="1"/>', count=1)
        backup("before update")
        book.save(RECORD + ".tmp")
        os.replace(RECORD + ".tmp", RECORD)
        print("Saved PnL Record.xlsx.")
    with open(LOG, "a", encoding="utf-8") as fh:
        fh.write("\n".join(log) + "\n")
    return 0


if __name__ == "__main__":
    sys.exit(main("--dry" in sys.argv))
