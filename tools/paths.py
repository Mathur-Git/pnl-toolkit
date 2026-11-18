"""Where everything lives. Every script imports this; nothing else hard-codes a path.

The toolkit folder is the parent of tools/. Copy start/, tools/ and README.md into an empty
folder and the scripts make the rest on first run.
"""
import glob
import os
import shutil
from datetime import datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RECORD = os.path.join(ROOT, "PnL Record.xlsx")
TEMPLATE = os.path.join(ROOT, "tools", "template", "PnL Record (blank).xlsx")
DAILY = os.path.join(ROOT, "Daily-statements")
DUPLICATES = os.path.join(DAILY, "duplicates")
MONTHLY = os.path.join(ROOT, "Monthly-statements")
EXPORTS = os.path.join(ROOT, "Outlook-export")
DATA = os.path.join(ROOT, "data")
BACKUPS = os.path.join(ROOT, "Backups")
KEEP_BACKUPS = 10

for d in (DAILY, MONTHLY, EXPORTS, DATA, BACKUPS):
    os.makedirs(d, exist_ok=True)


def backup(tag):
    """Copy the record into Backups/ before a save; keep only the newest KEEP_BACKUPS."""
    if not os.path.exists(RECORD):
        return None
    dest = os.path.join(BACKUPS, f"PnL Record {datetime.now():%Y-%m-%d %H%M%S} {tag}.xlsx")
    shutil.copy2(RECORD, dest)
    for old in sorted(glob.glob(os.path.join(BACKUPS, "*.xlsx")), key=os.path.getmtime)[:-KEEP_BACKUPS]:
        os.remove(old)
    return dest


def record_is_open():
    if not os.path.exists(RECORD):
        return False
    try:
        open(RECORD, "r+b").close()
        return False
    except PermissionError:
        return True
