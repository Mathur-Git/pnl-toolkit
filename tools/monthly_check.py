"""Reconcile the daily statements against Wedbush's monthly statements, from the ledger.

For each monthly statement and each account on it:
  month-end account value at market  vs  the daily statement's figure for that date
  the month's TC (TOTAL COMMISSION AND FEES)  vs  the sum of the daily TC for that month
check() returns one result per month and account; update_record prints only what is new or
unexplained.

carry_ins(): TC the daily statements lack (an account missing from them for part of the year),
taken from the monthly statements, so the journal's TC is complete:
  within a month: the monthly TC less the daily sum;
  before an account's first monthly statement: its year-to-date commission and fees less that
  month's TC.
Both dated that month end: the same row as the account's first P&L, which (the account value being
cumulative) also carries everything before it. Zero when the daily statements cover everything.
"""
import os
import sys
from datetime import date, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ledger  # noqa: E402


def carry_ins(roles=None, days=None, monthly=None):
    roles = roles or ledger.roles()
    days = days or ledger.daily_rows()
    monthly = monthly or ledger.monthly_rows()
    tcol = {"rates": "tc_rates", "vix": "tc_vix"}
    dates = {d["date"] for d in days}
    out, seen = [], set()
    for end in sorted({e for e, _ in monthly}):
        for role, acct in roles.items():
            m = monthly.get((end, acct))
            if not m or end not in dates:
                continue
            mtc = -m.get("TOTAL COMMISSION AND FEES", 0.0)
            dsum = sum(d[tcol[role]] or 0 for d in days if d["date"][:7] == end[:7])
            amount, why = 0.0, []
            if role not in seen:
                seen.add(role)
                ytd = -(m.get("YTD COMMISSIONS", 0.0) + m.get("YTD FEES", 0.0))
                if ytd - mtc > 0.005:
                    amount += ytd - mtc
                    why.append(f"before {end[:7]}: year to date {ytd:,.2f} less the month {mtc:,.2f}")
            if mtc - dsum > 0.005:
                amount += mtc - dsum
                why.append(f"in {end[:7]}: month {mtc:,.2f} less daily {dsum:,.2f}")
            if amount > 0.005:
                out.append({"date": end, "role": role, "amount": round(amount, 2),
                            "why": f"{role} TC missing from the daily statements, from the {end} monthly "
                                   f"statement ({'; '.join(why)})"})
    return out


def check():
    roles, days, monthly = ledger.roles(), ledger.daily_rows(), ledger.monthly_rows()
    by_date = {d["date"]: d for d in days}
    cols = {"rates": ("main_acct_value", "tc_rates"), "vix": ("vix_acct_value", "tc_vix")}
    filled = {(c["date"][:7], c["role"]) for c in carry_ins(roles, days, monthly)}
    out = []
    for end in sorted({e for e, _ in monthly}):
        for role, acct in roles.items():
            m = monthly.get((end, acct))
            if not m:
                continue
            vcol, tcol = cols[role]
            mv = m.get("ACCOUNT VALUE AT MARKET")
            mtc = -m.get("TOTAL COMMISSION AND FEES", 0.0)
            dv = by_date.get(end, {}).get(vcol)
            in_month = [d for d in days if d["date"][:7] == end[:7]]
            dtc = round(sum(d[tcol] or 0 for d in in_month), 2)
            first = date(int(end[:4]), int(end[5:7]), 1)
            weekdays = sum(1 for k in range(int(end[8:])) if (first + timedelta(k)).weekday() < 5)
            have = sum(1 for d in in_month if d[vcol] is not None)
            ok_v = dv is not None and mv is not None and abs(dv - mv) < 0.005
            ok_tc = abs(dtc - mtc) < 0.005
            notes = ([] if ok_v else ["value differs" if dv is not None else "no daily figure"]) + \
                    ([] if ok_tc else [f"TC off by {dtc - mtc:+,.2f}"])
            if not ok_tc and have < weekdays:
                notes.append(f"daily statements cover {have} of {weekdays} weekdays")
            if not ok_tc and (end[:7], role) in filled:
                notes.append("gap filled in the journal from this monthly statement")
            out.append({"month_end": end, "account": role, "ok": ok_v and ok_tc,
                        "known": ok_v and (end[:7], role) in filled,
                        "result": "agrees" if ok_v and ok_tc else "; ".join(notes)})
    return out


def main():
    res = check()
    for r in res:
        print(f"  {r['month_end']} {r['account']:<6} {r['result']}")
    print(f"  {sum(r['ok'] for r in res)} of {len(res)} month/account checks agree.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
