"""Make tools/template/PnL Record (blank).xlsx from the current record: same sheets, formulas and
style, every statement figure and typed value removed (journal D:S, the Statement's adjustment),
author fields cleared. update_record.py starts a new folder's record from it. Re-run after changing
the record's layout or style.
"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from paths import RECORD, TEMPLATE  # noqa: E402
from xlsxpatch import Book  # noqa: E402


def main():
    book = Book(RECORD)
    parts = book.sheet_parts()
    j = book.sheet(parts["IMR+PnL"], "IMR+PnL")
    j.expand_shared()
    for r in range(4, max(j.rows) + 1):  # every date after C3 rolls forward from the row above
        if j.raw(f"C{r}"):
            j.set_formula(f"B{r}", f'TEXT(C{r},"DDD")')
            j.set_formula(f"C{r}", f'C{r - 1}+IF($B{r - 1}="Fri",3,1)')
    for r in range(3, max(j.rows) + 1):
        for c in "DEFGHIJKLMNOPQRS":
            if j.raw(f"{c}{r}"):
                j.blank(f"{c}{r}")
    for r in j.rows:  # no collapsed row groups in a new record
        j.rows[r][0] = re.sub(r'\s*(?:hidden|outlineLevel|collapsed)="[^"]*"', "", j.rows[r][0])
    j.head = re.sub(r'\s*outlineLevelRow="\d+"', "", j.head)
    j.head = re.sub(r'topLeftCell="[A-Z]+\d+"', 'topLeftCell="A3"', j.head)
    s = book.sheet(parts["Statement"], "Statement")
    s.blank("C14")  # the adjustment: the one typed figure on the Statement
    for sheet in (j, s):  # formulas keep no calculated values: no figure of this record travels
        for r in sheet.rows:
            for c, raw in list(sheet.rows[r][1].items()):
                if "<f" in raw:
                    raw = re.sub(r"<v>.*?</v>", "", raw, flags=re.S)
                    sheet.rows[r][1][c] = re.sub(r'\st="(?:str|n|e|b)"', "", raw)
    core = "docProps/core.xml"
    if core in book.parts:
        book.sub(core, r"<dc:creator>.*?</dc:creator>", "<dc:creator></dc:creator>")
        book.sub(core, r"<cp:lastModifiedBy>.*?</cp:lastModifiedBy>", "<cp:lastModifiedBy></cp:lastModifiedBy>")
    # a typed note no cell uses any more still sits in the shared strings: blank every unused one
    # (kept in place, so the indices the cells use do not move)
    sst = "xl/sharedStrings.xml"
    if sst in book.parts:
        used = {int(v) for sheet in (j, s) for r in sheet.rows for raw in sheet.rows[r][1].values()
                for v in re.findall(r'\st="s"[^>]*>(?:<f[^<]*</f>)?<v>(\d+)</v>', raw)}
        n = iter(range(10 ** 9))
        book.sub(sst, r"<si>.*?</si>", lambda m: m.group(0) if next(n) in used else "<si><t></t></si>")
    os.makedirs(os.path.dirname(TEMPLATE), exist_ok=True)
    book.save(TEMPLATE)
    print(f"Wrote {TEMPLATE}")


if __name__ == "__main__":
    sys.exit(main())
