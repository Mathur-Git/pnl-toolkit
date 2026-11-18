"""The statement ledger: data/statements.db (SQLite). Every Wedbush PDF is read once, when it
first arrives, and everything in it is kept here; nothing reads a PDF again. If the PDFs are lost,
the ledger stands on its own. Open it with any SQLite browser, or from Python:

    import sqlite3, pandas as pd
    con = sqlite3.connect("data/statements.db")
    pd.read_sql("SELECT * FROM fills WHERE date = '2026-11-13'", con)

Tables (dates ISO, amounts as the statement states them: costs and losses negative):
    files         every PDF read: file, kind (daily/monthly), date, sha256, status, detail, used
                  (0 = a second copy of a day that was not used)
    accounts      account, role (rates = trades SR3/SR1/ZQ, vix = trades VX)
    summary       per day and account, each line of the account summary (BEGINNING BALANCE,
                  COMMISSION, EXCHANGE FEE, ..., ACCOUNT VALUE AT MARKET, INITIAL MARGIN
                  REQUIREMENT, ...) and the M-T-D / Y-T-D lines (MTD REALIZED, YTD FEES, ...)
    fill_blocks   per day, account and contract: contracts bought and sold, average prices,
                  commission, exchange fee, clearing fee, other fees, total
    fills         every fill: trade date, product, contract, side (B/S), qty, price, spread leg
    realized      every closing line of "purchase & sale": product, contract, qty, gross P&L
    positions     open positions at the close: product, contract, side (L/S), qty, average
                  price, settle, open P&L
    adjustments   "trade related adjustments": description, fee line, amount (+ = refund)
    monthly       per monthly statement and account, the same lines as summary

Checks before a day is stored (a failure keeps it out and says why):
    each fill block: its parts = its TOTAL COMM AND FEE; its fills' bought / sold = its n* counts
    each account: blocks + adjustments = TOTAL COMMISSION AND FEES; realized lines = GROSS PROFIT
      OR LOSS; open positions = OPEN TRADE EQUITY; beginning + fees + gross = ending balance;
      ending + open trade equity = total equity = account value at market (a figure read off the
      wrong row breaks this); the three currency columns agree
    the date inside the PDF = the date in its file name
"""
import hashlib
import json
import os
import re
import sqlite3
import sys
from datetime import date, datetime

import pdfplumber

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from paths import DATA, DAILY, DUPLICATES, MONTHLY  # noqa: E402

DB = os.path.join(DATA, "statements.db")
PRODUCTS = {"CME 3MO SOFR": "SR3", "CME 1MO SOFR": "SR1", "CBT 30 DAY IR": "ZQ", "CFE VIX": "VX"}
PRODUCT_ORDER = ["SR3", "SR1", "ZQ", "VX"]
RATES = {"SR3", "SR1", "ZQ"}
MONTHS = "JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC"
AMT = r"(-?[\d,]*\.\d\d-?)"
RE_LOT = re.compile(rf"^(\d{{1,2}}/\d{{2}}/\d) US (\d+) ((?:{MONTHS}) \d\d) (.+?) (\S+) ([\d.]+) US(?: {AMT})?$")
RE_SUMLINE = re.compile(rf"^([A-Z][A-Z &/.]+?) {AMT} {AMT} {AMT}$")
RE_COST = re.compile(rf"(TOTAL COMM AND FEE|COMMISSION|[A-Z]+ FEE) US {AMT}$")
RE_AVG = re.compile(r"AVG (LONG|SHORT): ([\d.]+)")
RE_LTD = re.compile(r"LTD-\s?(\d{1,2}/\d{1,2}/\d\d)")
RE_POS = re.compile(rf"^(\d+)\* LTD-\s?(\d{{1,2}}/\d{{1,2}}/\d\d) CLOSE ([\d.]+) {AMT}$")
RE_REAL = re.compile(rf"^US (\d+)\* (\d+)\* LTD-\s?(\d{{1,2}}/\d{{1,2}}/\d\d) GROSS PROFIT OR LOSS US {AMT}$")
RE_FILE = re.compile(r"^(m?)statement(\d{4})(\d{2})(\d{2})(?:__\d+)?\.pdf$", re.I)

SCHEMA = """
CREATE TABLE IF NOT EXISTS files (file TEXT PRIMARY KEY, kind TEXT, date TEXT, sha256 TEXT,
    read_at TEXT, status TEXT, detail TEXT, used INTEGER);
CREATE TABLE IF NOT EXISTS accounts (account TEXT PRIMARY KEY, role TEXT);
CREATE TABLE IF NOT EXISTS summary (date TEXT, account TEXT, label TEXT, value REAL,
    PRIMARY KEY (date, account, label));
CREATE TABLE IF NOT EXISTS fill_blocks (date TEXT, account TEXT, block INTEGER, product TEXT,
    contract TEXT, last_trade TEXT, bought INTEGER, sold INTEGER, avg_bought REAL, avg_sold REAL,
    commission REAL, exchange_fee REAL, clearing_fee REAL, other_fees REAL, total REAL,
    PRIMARY KEY (date, account, block));
CREATE TABLE IF NOT EXISTS fills (date TEXT, account TEXT, block INTEGER, seq INTEGER,
    trade_date TEXT, product TEXT, contract TEXT, side TEXT, qty INTEGER, price REAL,
    spread INTEGER, PRIMARY KEY (date, account, block, seq));
CREATE TABLE IF NOT EXISTS realized (date TEXT, account TEXT, seq INTEGER, product TEXT,
    contract TEXT, last_trade TEXT, qty INTEGER, gross_pl REAL, PRIMARY KEY (date, account, seq));
CREATE TABLE IF NOT EXISTS positions (date TEXT, account TEXT, product TEXT, contract TEXT,
    last_trade TEXT, side TEXT, qty INTEGER, avg_price REAL, settle REAL, open_pl REAL,
    PRIMARY KEY (date, account, product, contract));
CREATE TABLE IF NOT EXISTS adjustments (date TEXT, account TEXT, seq INTEGER, description TEXT,
    label TEXT, amount REAL, PRIMARY KEY (date, account, seq));
CREATE TABLE IF NOT EXISTS monthly (month_end TEXT, account TEXT, label TEXT, value REAL,
    PRIMARY KEY (month_end, account, label));
"""
DAY_TABLES = ("summary", "fill_blocks", "fills", "realized", "positions", "adjustments")


def num(s):
    if s is None:
        return None
    s = s.replace(",", "")
    return -float(s[:-1]) if s.endswith("-") else float(s)


def product_of(desc):
    return next((code for key, code in PRODUCTS.items() if desc.startswith(key)), desc)


def iso_mdy(s, year_hint):
    """'11/13/6' or '12/19/28' -> ISO. A one-digit year is the decade's year nearest year_hint."""
    m, d, y = (int(p) for p in s.split("/"))
    if y < 10:
        y = year_hint - ((year_hint - y) % 10)
    elif y < 100:
        y += 2000
    return date(y, m, d).isoformat()


def lines_of(path):
    """Each text line of the PDF with its words' positions: [(text, words, sell_x)] where sell_x is
    the left edge of that page's SELL column header (None if the page has none)."""
    out = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            sell_x = next((w["x0"] for w in words if w["text"] == "SELL"), None)
            rows = []
            for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
                if rows and abs(rows[-1][0] - w["top"]) <= 2.0:
                    rows[-1][1].append(w)
                else:
                    rows.append([w["top"], [w]])
            for _, ws in rows:
                ws.sort(key=lambda w: w["x0"])
                out.append((" ".join(w["text"] for w in ws), ws, sell_x))
    return out


def parse_daily(path):
    """Everything one daily statement says, by account, plus the problems its checks found."""
    lines = lines_of(path)
    texts = [t for t, _, _ in lines]
    if not any("D A I L Y S T A T E M E N T" in t for t in texts):
        raise ValueError("not a daily statement")
    stmt_date = None
    for t in texts:
        m = re.search(r"STATEMENT DATE: ([A-Z]{3}) (\d{1,2}), (\d{4})", t)
        if m:
            stmt_date = datetime.strptime(" ".join(m.groups()), "%b %d %Y").date()
            break
    if stmt_date is None:
        raise ValueError("no statement date")
    yr = stmt_date.year

    acct, section = None, None
    d = {"summary": {}, "blocks": [], "fills": [], "realized": [], "positions": [], "adjustments": []}
    problems = []
    blk = None  # the confirmation block being read
    last_lot = None  # (product, contract) of the last lot line, for P&S and positions summaries
    adj_desc = None

    def new_block():
        return {"account": acct, "fills": [], "bought": None, "sold": None, "avg_b": None, "avg_s": None,
                "parts": {}, "total": None, "ltd": None}

    for i, (t, ws, sell_x) in enumerate(lines):
        m = re.search(r"ACCOUNT NUMBER: P ([A-Z0-9]+)", t)
        if m:
            acct = m.group(1)
            continue
        if "C O N F I R M A T I O N" in t:
            section, blk = "confirm", None
            continue
        if "P U R C H A S E" in t:
            section = "ps"
            continue
        if "O P E N P O S I T I O N S" in t:
            section = "pos"
            continue
        if "TRADE RELATED ADJUSTMENTS" in t:
            section, adj_desc = "adj", None
            continue
        if "SEGREGATED" in t:
            section = "summary"
            continue
        if t.startswith("SUMMARY") and "M-T-D" in t:
            section = "ytd"
            continue

        if section == "summary":
            m = RE_SUMLINE.match(t)
            if m and acct:
                vals = [num(v) for v in m.groups()[1:]]
                if max(vals) - min(vals) > 0.005:
                    problems.append(f"{acct} {m.group(1)}: the three columns differ")
                d["summary"].setdefault(acct, {})[m.group(1).strip()] = vals[-1]
            continue
        if section == "ytd":
            if t in ("REALIZED PROFIT & LOSS", "COMMISSIONS", "FEES") and i + 1 < len(lines):
                v = lines[i + 1][0].split()
                if len(v) == 3 and v[0] == "USD":
                    key = {"REALIZED PROFIT & LOSS": "REALIZED", "COMMISSIONS": "COMMISSIONS",
                           "FEES": "FEES"}[t]
                    d["summary"].setdefault(acct, {})[f"MTD {key}"] = num(v[1])
                    d["summary"][acct][f"YTD {key}"] = num(v[2])
            continue

        lot = RE_LOT.match(t)
        if lot:
            last_lot = (product_of(lot.group(4)), lot.group(3))
        if section == "confirm":
            if lot:
                if blk is None or blk["total"] is not None:
                    blk = new_block()
                    d["blocks"].append(blk)
                qty_word = next(w for w in ws if w["text"] == lot.group(2) and w["x0"] > 60)
                side = "S" if sell_x is not None and qty_word["x1"] > sell_x else "B"
                blk["fills"].append({"trade_date": iso_mdy(lot.group(1), yr), "product": last_lot[0],
                                     "contract": last_lot[1], "side": side, "qty": int(lot.group(2)),
                                     "price": float(lot.group(6)), "spread": 0})
                continue
            if t == "SPREAD TRADE" and blk and blk["fills"]:
                blk["fills"][-1]["spread"] = 1
                continue
            if blk is None:
                continue
            if "LTD-" in t:
                m = RE_LTD.search(t)
                blk["ltd"] = iso_mdy(m.group(1), yr) if m else None
                b = s = 0
                for w in ws:
                    if re.fullmatch(r"\d+\*", w["text"]):
                        if sell_x is not None and w["x1"] > sell_x:
                            s += int(w["text"][:-1])
                        else:
                            b += int(w["text"][:-1])
                blk["bought"], blk["sold"] = b, s
            for side, px in RE_AVG.findall(t):
                blk["avg_b" if side == "LONG" else "avg_s"] = float(px)
            m = RE_COST.search(t)
            if m:
                if m.group(1) == "TOTAL COMM AND FEE":
                    blk["total"] = num(m.group(2))
                else:
                    blk["parts"][m.group(1)] = blk["parts"].get(m.group(1), 0.0) + num(m.group(2))
            continue
        if section == "ps":
            m = RE_REAL.match(t)
            if m and last_lot:
                d["realized"].append({"account": acct, "product": last_lot[0], "contract": last_lot[1],
                                      "last_trade": iso_mdy(m.group(3), yr), "qty": int(m.group(1)),
                                      "gross_pl": num(m.group(4))})
            continue
        if section == "pos":
            m = RE_POS.match(t)
            if m and last_lot:
                side = None
                if i + 1 < len(lines):
                    a = RE_AVG.search(lines[i + 1][0])
                    side = {"LONG": "L", "SHORT": "S"}[a.group(1)] if a else None
                    avg = float(a.group(2)) if a else None
                d["positions"].append({"account": acct, "product": last_lot[0], "contract": last_lot[1],
                                       "last_trade": iso_mdy(m.group(2), yr), "side": side,
                                       "qty": int(m.group(1)), "avg_price": avg,
                                       "settle": float(m.group(3)), "open_pl": num(m.group(4))})
            continue
        if section == "adj":
            m = RE_COST.search(t)
            if m:
                if m.group(1) != "TOTAL COMM AND FEE":
                    d["adjustments"].append({"account": acct, "description": adj_desc or "",
                                             "label": m.group(1), "amount": num(m.group(2))})
            elif t.startswith("US ") and adj_desc is None:
                adj_desc = re.sub(r"\s+US [\d,.]+-?$", "", t[3:]).strip()
            continue

    # ---- checks
    for b in d["blocks"]:
        parts = sum(b["parts"].values())
        prod = b["fills"][0]["product"] if b["fills"] else "?"
        if b["total"] is None or abs(parts - b["total"]) > 0.005:
            problems.append(f"{prod} block: parts {parts:.2f} vs total {b['total']}")
        fb = sum(f["qty"] for f in b["fills"] if f["side"] == "B")
        fs = sum(f["qty"] for f in b["fills"] if f["side"] == "S")
        if (fb, fs) != (b["bought"], b["sold"]):
            problems.append(f"{prod} block: fills bought/sold {fb}/{fs} vs stated {b['bought']}/{b['sold']}")
        # sides checked a second, independent way: each side's average fill price must be the
        # statement's own AVG LONG / AVG SHORT (and a side with no average has no fills)
        for side, qty, avg in (("B", fb, b["avg_b"]), ("S", fs, b["avg_s"])):
            if qty == 0 and avg is None:
                continue
            got = sum(f["qty"] * f["price"] for f in b["fills"] if f["side"] == side) / qty if qty else None
            if avg is None or got is None or abs(got - avg) > 0.00002:
                problems.append(f"{prod} block: average {'bought' if side == 'B' else 'sold'} {got} vs stated {avg}")
    for a, s in d["summary"].items():
        tc = sum(b["total"] or 0 for b in d["blocks"] if b["account"] == a)
        tc += sum(x["amount"] for x in d["adjustments"] if x["account"] == a)
        if abs(tc - s.get("TOTAL COMMISSION AND FEES", 0.0)) > 0.005:
            problems.append(f"{a}: TC {tc:.2f} vs statement {s.get('TOTAL COMMISSION AND FEES', 0.0):.2f}")
        gross = sum(x["gross_pl"] for x in d["realized"] if x["account"] == a)
        if abs(gross - s.get("GROSS PROFIT OR LOSS", 0.0)) > 0.005:
            problems.append(f"{a}: realized {gross:.2f} vs gross P&L {s.get('GROSS PROFIT OR LOSS', 0.0):.2f}")
        ote = sum(x["open_pl"] for x in d["positions"] if x["account"] == a)
        if abs(ote - s.get("OPEN TRADE EQUITY", 0.0)) > 0.005:
            problems.append(f"{a}: positions {ote:.2f} vs open trade equity {s.get('OPEN TRADE EQUITY', 0.0):.2f}")
        if "ENDING BALANCE" in s:
            flow = s.get("BEGINNING BALANCE", 0.0) + s.get("TOTAL COMMISSION AND FEES", 0.0) + s.get("GROSS PROFIT OR LOSS", 0.0)
            if abs(flow - s["ENDING BALANCE"]) > 0.005:
                problems.append(f"{a}: beginning + fees + gross {flow:.2f} vs ending {s['ENDING BALANCE']:.2f}")
        if "ACCOUNT VALUE AT MARKET" in s:
            te = s.get("ENDING BALANCE", 0.0) + s.get("OPEN TRADE EQUITY", 0.0)
            if abs(te - s.get("TOTAL EQUITY", te)) > 0.005 or abs(te - s["ACCOUNT VALUE AT MARKET"]) > 0.005:
                problems.append(f"{a}: ending + open trade equity {te:.2f} vs value at market {s['ACCOUNT VALUE AT MARKET']:.2f}")
    m = RE_FILE.match(os.path.basename(path))
    if m and "-".join(m.groups()[1:]) != stmt_date.isoformat():
        problems.append(f"file name says {'-'.join(m.groups()[1:])}, statement says {stmt_date}")
    for b in d["blocks"]:
        for f in b["fills"]:
            if f["product"] not in PRODUCT_ORDER:
                problems.append(f"unknown product '{f['product']}'")
    d["date"], d["problems"] = stmt_date.isoformat(), problems
    return d


def parse_monthly(path):
    """Month-end figures per account: the account summary lines and the M-T-D / Y-T-D lines."""
    lines = lines_of(path)
    out, acct, section = {}, None, None
    for i, (t, _, _) in enumerate(lines):
        m = re.search(r"ACCOUNT NUMBER: P ([A-Z0-9]+)", t)
        if m:
            acct = m.group(1)
        if "SEGREGATED" in t:
            section = "summary"
            continue
        if t.startswith("SUMMARY") and "M-T-D" in t:
            section = "ytd"
            continue
        if section == "summary":
            m = RE_SUMLINE.match(t)
            if m and acct:
                out.setdefault(acct, {})[m.group(1).strip()] = num(m.group(4))
        if section == "ytd" and t in ("REALIZED P&L", "REALIZED PROFIT & LOSS", "COMMISSION",
                                      "COMMISSIONS", "FEES") and i + 1 < len(lines):
            v = lines[i + 1][0].split()
            if len(v) == 3 and v[0] == "USD" and acct:
                key = "REALIZED" if t.startswith("REALIZED") else ("FEES" if t == "FEES" else "COMMISSIONS")
                out.setdefault(acct, {})[f"MTD {key}"] = num(v[1])
                out[acct][f"YTD {key}"] = num(v[2])
    return out


# ------------------------------------------------------------------ the database

def connect():
    con = sqlite3.connect(DB)
    con.executescript(SCHEMA)
    return con


def sha(path):
    with open(path, "rb") as fh:
        return hashlib.sha256(fh.read()).hexdigest()


def store_day(con, d):
    day = d["date"]
    for t in DAY_TABLES:
        con.execute(f"DELETE FROM {t} WHERE date = ?", (day,))
    for a, s in d["summary"].items():
        con.executemany("INSERT INTO summary VALUES (?,?,?,?)", [(day, a, k, v) for k, v in s.items()])
    for n, b in enumerate(d["blocks"], 1):
        p = b["parts"]
        other = sum(v for k, v in p.items() if k not in ("COMMISSION", "EXCHANGE FEE", "CLEARING FEE"))
        f0 = b["fills"][0] if b["fills"] else {"product": None, "contract": None}
        con.execute("INSERT INTO fill_blocks VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    (day, b["account"], n, f0["product"], f0["contract"], b["ltd"], b["bought"], b["sold"],
                     b["avg_b"], b["avg_s"], p.get("COMMISSION", 0.0), p.get("EXCHANGE FEE", 0.0),
                     p.get("CLEARING FEE", 0.0), other, b["total"]))
        con.executemany("INSERT INTO fills VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                        [(day, b["account"], n, k, f["trade_date"], f["product"], f["contract"], f["side"],
                          f["qty"], f["price"], f["spread"]) for k, f in enumerate(b["fills"], 1)])
    con.executemany("INSERT INTO realized VALUES (?,?,?,?,?,?,?,?)",
                    [(day, x["account"], k, x["product"], x["contract"], x["last_trade"], x["qty"], x["gross_pl"])
                     for k, x in enumerate(d["realized"], 1)])
    con.executemany("INSERT OR REPLACE INTO positions VALUES (?,?,?,?,?,?,?,?,?,?)",
                    [(day, x["account"], x["product"], x["contract"], x["last_trade"], x["side"], x["qty"],
                      x["avg_price"], x["settle"], x["open_pl"]) for x in d["positions"]])
    con.executemany("INSERT INTO adjustments VALUES (?,?,?,?,?,?)",
                    [(day, x["account"], k, x["description"], x["label"], x["amount"])
                     for k, x in enumerate(d["adjustments"], 1)])


def states_margin(d, con):
    """Does this copy state the rates account's initial margin? (one 9 Jan 2026 copy does not)"""
    rates = dict(con.execute("SELECT role, account FROM accounts").fetchall()).get("rates")
    accts = [rates] if rates else list(d["summary"])
    return any("INITIAL MARGIN REQUIREMENT" in d["summary"].get(a, {}) for a in accts)


def assign_roles(con, parsed):
    roles = dict(con.execute("SELECT role, account FROM accounts").fetchall())
    for d in parsed:
        seen = {}
        for b in d["blocks"]:
            for f in b["fills"]:
                seen.setdefault(b["account"], set()).add(f["product"])
        for x in d["positions"]:
            seen.setdefault(x["account"], set()).add(x["product"])
        for a, prods in seen.items():
            if prods & RATES and "rates" not in roles:
                roles["rates"] = a
            if "VX" in prods and "vix" not in roles:
                roles["vix"] = a
    con.executemany("INSERT OR REPLACE INTO accounts VALUES (?,?)", [(a, r) for r, a in roles.items()])
    return roles


def update(log=print):
    """Read every PDF the ledger has not seen. Returns a dict of what happened, for the run report."""
    con = connect()
    known = {r[0]: r[1] for r in con.execute("SELECT sha256, file FROM files")}
    daily = sorted(f for f in os.listdir(DAILY) if RE_FILE.match(f) and not f.lower().startswith("m"))
    dups = sorted(f for f in os.listdir(DUPLICATES)) if os.path.isdir(DUPLICATES) else []
    monthly = sorted(f for f in os.listdir(MONTHLY) if RE_FILE.match(f) and f.lower().startswith("m"))
    report = {"new_days": [], "failed": [], "replaced": [], "new_months": [], "same_file": []}
    now = datetime.now().isoformat(timespec="seconds")

    todo = [(os.path.join(DAILY, f), f) for f in daily] + [(os.path.join(DUPLICATES, f), f) for f in dups
                                                           if RE_FILE.match(f)]
    parsed = []
    for path, f in todo:
        h = sha(path)
        if h in known:
            continue
        try:
            d = parse_daily(path)
        except Exception as e:  # noqa: BLE001
            con.execute("INSERT OR REPLACE INTO files VALUES (?,?,?,?,?,?,?,?)",
                        (f, "daily", None, h, now, "failed", str(e), 0))
            report["failed"].append((f, str(e)))
            continue
        d["file"], d["sha"] = f, h
        parsed.append(d)
    assign_roles(con, parsed)

    for d in sorted(parsed, key=lambda d: (d["date"], d["file"])):
        if d["problems"]:
            con.execute("INSERT OR REPLACE INTO files VALUES (?,?,?,?,?,?,?,?)",
                        (d["file"], "daily", d["date"], d["sha"], now, "failed", "; ".join(d["problems"]), 0))
            report["failed"].append((d["file"], "; ".join(d["problems"])))
            continue
        held = con.execute("SELECT file FROM files WHERE kind='daily' AND date=? AND used=1", (d["date"],)).fetchone()
        use = True
        if held:  # a second copy of a day already stored: the copy that states the margin wins
            old_avm = dict(con.execute("SELECT account, value FROM summary WHERE date=? AND label='ACCOUNT VALUE AT MARKET'",
                                       (d["date"],)).fetchall())
            new_avm = {a: s.get("ACCOUNT VALUE AT MARKET") for a, s in d["summary"].items() if "ACCOUNT VALUE AT MARKET" in s}
            if old_avm != new_avm:
                con.execute("INSERT OR REPLACE INTO files VALUES (?,?,?,?,?,?,?,?)",
                            (d["file"], "daily", d["date"], d["sha"], now, "failed",
                             f"second copy disagrees with {held[0]} on account value", 0))
                report["failed"].append((d["file"], f"second copy of {d['date']} disagrees with {held[0]}"))
                continue
            old_margin = con.execute("SELECT 1 FROM summary s JOIN accounts a ON a.account=s.account AND a.role='rates' "
                                     "WHERE s.date=? AND s.label='INITIAL MARGIN REQUIREMENT'", (d["date"],)).fetchone()
            use = states_margin(d, con) and not old_margin
            if use:
                con.execute("UPDATE files SET used=0, detail=? WHERE file=?",
                            (f"second copy {d['file']} used instead: it states the margin", held[0]))
                report["replaced"].append((d["date"], d["file"], held[0]))
        if use:
            store_day(con, d)
            if not held:
                report["new_days"].append(d["date"])
        con.execute("INSERT OR REPLACE INTO files VALUES (?,?,?,?,?,?,?,?)",
                    (d["file"], "daily", d["date"], d["sha"], now, "ok",
                     "" if use else f"second copy; {held[0]} kept", 1 if use else 0))

    for f in monthly:
        path = os.path.join(MONTHLY, f)
        h = sha(path)
        if h in known:
            continue
        m = RE_FILE.match(f)
        end = "-".join(m.groups()[1:])
        try:
            data = parse_monthly(path)
        except Exception as e:  # noqa: BLE001
            report["failed"].append((f, str(e)))
            continue
        con.execute("DELETE FROM monthly WHERE month_end=?", (end,))
        for a, s in data.items():
            con.executemany("INSERT INTO monthly VALUES (?,?,?,?)", [(end, a, k, v) for k, v in s.items()])
        con.execute("INSERT OR REPLACE INTO files VALUES (?,?,?,?,?,?,?,?)",
                    (f, "monthly", end, h, now, "ok", "", 1))
        report["new_months"].append(end)
    con.commit()
    con.close()
    return report


# ------------------------------------------------------------------ reading it back

def roles():
    con = connect()
    r = dict(con.execute("SELECT role, account FROM accounts").fetchall())
    con.close()
    return r


def daily_rows():
    """One dict a stored day, the figures the journal is written from."""
    con = connect()
    role = dict(con.execute("SELECT role, account FROM accounts").fetchall())
    ra, va = role.get("rates"), role.get("vix")
    out = []
    for (day,) in con.execute("SELECT DISTINCT date FROM summary ORDER BY date").fetchall():
        s = {}
        for a, label, v in con.execute("SELECT account, label, value FROM summary WHERE date=?", (day,)):
            s.setdefault(a, {})[label] = v
        tc = dict(con.execute("SELECT product, -SUM(total) FROM fill_blocks WHERE date=? GROUP BY product", (day,)).fetchall())
        vol = dict(con.execute("SELECT product, SUM(qty) FROM fills WHERE date=? GROUP BY product", (day,)).fetchall())
        rates, vix = s.get(ra, {}), s.get(va, {})
        row = {"date": day,
               "main_acct_value": rates.get("ACCOUNT VALUE AT MARKET"),
               "main_init_margin": rates.get("INITIAL MARGIN REQUIREMENT", 0.0),
               "vix_acct_value": vix.get("ACCOUNT VALUE AT MARKET"),
               "vix_init_margin": vix.get("INITIAL MARGIN REQUIREMENT", 0.0) if vix else None,
               "tc_rates": round(-rates.get("TOTAL COMMISSION AND FEES", 0.0), 2),
               "tc_vix": round(-vix.get("TOTAL COMMISSION AND FEES", 0.0), 2)}
        for p in PRODUCT_ORDER:
            row[f"cost_{p}"] = round(tc.get(p) or 0.0, 2)
            row[f"vol_{p}"] = int(vol.get(p) or 0)
        out.append(row)
    con.close()
    return out


def monthly_rows():
    """{(month_end, account): {label: value}} from the monthly statements."""
    con = connect()
    out = {}
    for end, a, label, v in con.execute("SELECT month_end, account, label, value FROM monthly"):
        out.setdefault((end, a), {})[label] = v
    con.close()
    return out


if __name__ == "__main__":
    print(json.dumps(update(), indent=1))
