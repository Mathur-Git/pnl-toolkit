"""Pull the statement PDFs out of an Outlook export (.pst).

Looks for the newest .pst in Outlook-export/ (or the toolkit folder itself, then moves it into
Outlook-export/). Only ever ADDS files; nothing is removed or overwritten:
  statementYYYYMMDD.pdf   -> Daily-statements/   (already there, same bytes: skipped;
                             same name, different bytes: Daily-statements/duplicates/)
  mstatementYYYYMMDD.pdf  -> Monthly-statements/ (same rules)
  anything else (notices) -> not saved, counted

The .pst is read by tools/wedbush-extract/pstlib.py + ltp.py: pure Python, no Outlook, no MAPI,
tested on a 179-email export in Sept 2026 and on a 238-email export on 18 Nov 2026. Quit Outlook
before running: it locks the file's header while open.
"""
import glob
import os
import re
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "wedbush-extract"))
import ltp  # noqa: E402
import pstlib  # noqa: E402
from paths import DAILY, DUPLICATES, EXPORTS, MONTHLY, ROOT  # noqa: E402

KINDS = [(re.compile(r"^statement\d{8}\.pdf$", re.I), DAILY, "daily"),
         (re.compile(r"^mstatement\d{8}\.pdf$", re.I), MONTHLY, "monthly")]


def find_pst():
    for loose in glob.glob(os.path.join(ROOT, "*.pst")):  # saved to the top folder: tidy it away
        dest = os.path.join(EXPORTS, os.path.basename(loose))
        if not os.path.exists(dest):
            shutil.move(loose, dest)
    found = sorted(glob.glob(os.path.join(EXPORTS, "*.pst")), key=os.path.getmtime)
    return found[-1] if found else None


def save(name, data, folder):
    """'added', 'same' or the duplicate's file name."""
    dest = os.path.join(folder, name)
    if os.path.exists(dest):
        with open(dest, "rb") as fh:
            if fh.read() == data:
                return "same"
        ddir = DUPLICATES if folder == DAILY else os.path.join(folder, "duplicates")
        os.makedirs(ddir, exist_ok=True)
        for kept in glob.glob(os.path.join(ddir, f"{name[:-4]}__*.pdf")):  # this copy already kept
            with open(kept, "rb") as fh:
                if fh.read() == data:
                    return "same"
        k = 2
        while os.path.exists(os.path.join(ddir, f"{name[:-4]}__{k}.pdf")):
            k += 1
        dest = os.path.join(ddir, f"{name[:-4]}__{k}.pdf")
        with open(dest, "wb") as fh:
            fh.write(data)
        return os.path.basename(dest)
    with open(dest, "wb") as fh:
        fh.write(data)
    return "added"


def main():
    pst = find_pst()
    if not pst:
        print("No Outlook export (.pst) found. Save it into the Outlook-export folder (README, step 1).")
        return 1
    try:
        p = pstlib.PST(pst)
    except PermissionError:
        print("The .pst is locked. Quit Outlook (classic) completely and run again.")
        return 1
    tally = {k: {"added": [], "same": 0, "dups": []} for _, _, k in KINDS}
    other = 0
    msgs = sorted(n for n in p.nbt if pstlib.nid_type(n) == pstlib.NID_TYPE_NORMAL_MESSAGE)
    for nid in msgs:
        pc = ltp.PropertyContext(p, nid)
        for snid, (bd, bs) in sorted(pc.subs.items()):
            if pstlib.nid_type(snid) != 0x05:  # not an attachment
                continue
            apc = ltp.PropertyContext(p, snid, bids=(bd, bs))
            name = os.path.basename(apc.get(0x3707) or apc.get(0x3704) or "").lower()
            data = apc.get(0x3701)
            if not isinstance(data, (bytes, bytearray)) or bytes(data[:5]) != b"%PDF-":
                continue
            kind = next(((folder, k) for rx, folder, k in KINDS if rx.match(name)), None)
            if not kind:
                other += 1
                continue
            res = save(name, bytes(data), kind[0])
            t = tally[kind[1]]
            if res == "added":
                t["added"].append(name)
            elif res == "same":
                t["same"] += 1
            else:
                t["dups"].append(res)

    d, m = tally["daily"], tally["monthly"]
    print(f"Outlook export: {len(msgs)} emails; {len(d['added'])} new daily and {len(m['added'])} new monthly "
          f"statement(s) saved.")
    for k, x in tally.items():
        if x["dups"]:
            print(f"  A second {k} copy with different contents, kept in duplicates: {', '.join(x['dups'])}")
    if not msgs:
        print("  The export holds no emails: check the right Outlook folder was exported.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
