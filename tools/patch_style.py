"""House style for PnL Record.xlsx, matching Final Charts (excel/CLAUDE.md §6a). Idempotent.

Aptos Narrow 11 throughout, title 11 bold navy; navy #1F4E79 header bars with white bold text,
all section bars one width (B:E); light-grey #BFBFBF hairline borders instead of black boxes;
P&L never red (a minus sign only); margins as whole numbers; costs with "-" for zero; no data
bars; navy tabs; zoom 100%. Values, formulas and the send block's own number formats untouched.

    python tools/patch_style.py [workbook]     (default: PnL Record.xlsx, copied to Backups/ first)
"""
import os
import re
import shutil
import sys
from datetime import datetime

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsxpatch import Book, Styles  # noqa: E402

NAVY = "FF1F4E79"
GREY = "FFBFBFBF"
MID = "FF808080"
F = '<name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/>'
FONTS = (f'<font><sz val="11"/><color theme="1"/>{F}</font>'
         f'<font><b/><sz val="11"/><color theme="1"/>{F}</font>'
         f'<font><b/><sz val="11"/><color rgb="{NAVY}"/>{F}</font>'          # title
         f'<font><b/><sz val="11"/><color theme="0"/>{F}</font>')           # on navy
HAIR = (f'<left style="thin"><color rgb="{GREY}"/></left><right style="thin"><color rgb="{GREY}"/>'
        f'</right><top style="thin"><color rgb="{GREY}"/></top><bottom style="thin"><color '
        f'rgb="{GREY}"/></bottom><diagonal/>')
BORDERS = ('<border><left/><right/><top/><bottom/><diagonal/></border>'
           + f'<border>{HAIR}</border>' * 3
           + f'<border><left/><right/><top/><bottom style="thin"><color rgb="{MID}"/></bottom>'
             '<diagonal/></border>'
           + f'<border><left/><right/><top style="thin"><color rgb="{MID}"/></top><bottom/>'
             '<diagonal/></border>')  # 5: the rule above a total row
PNL =r'#,##0.00;\-#,##0.00'               # never red: his call, 2026-11-18
COST = r'#,##0.00;\-#,##0.00;"-"'
VOL = r'#,##0;\-#,##0;"-"'                  # contracts: not a position, so no blue/red


FILLS = ('<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
         f'<fill><patternFill patternType="solid"><fgColor rgb="{NAVY}"/><bgColor indexed="64"/></patternFill></fill>'
         '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>')
BAR_XML = ('<xf numFmtId="0" fontId="3" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" '
           'applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf>')  # navy bar


def _items(blk, tag):
    return re.findall(rf"<{tag}(?:\s[^>]*)?/>|<{tag}(?:\s[^>]*)?>.*?</{tag}>", blk, re.S)


def house_first(st, tag, house, attr):
    """Put the house entries first in <tag>s and keep every other entry after them, remapping each
    xf's attr. Excel adds entries and renumbers them when someone formats and saves the record, so
    an index means nothing until this has run; overwriting the block instead broke the file
    (2026-11-18: two italic fonts Excel had added were dropped while cells still pointed at them)."""
    m = re.search(rf"(<{tag}s\b[^>]*>)(.*?)(</{tag}s>)", st, re.S)
    old, new = _items(m.group(2), tag), _items(house, tag)
    remap = {}
    for i, it in enumerate(old):
        if it not in new:
            new.append(it)
        remap[i] = new.index(it)
    head = re.sub(r'count="\d+"', f'count="{len(new)}"', m.group(1))
    st = st[:m.start()] + head + "".join(new) + m.group(3) + st[m.end():]
    fix = lambda x: re.sub(rf'\b{attr}="(\d+)"', lambda a: f'{attr}="{remap.get(int(a.group(1)), 0)}"', x.group(0))  # noqa: E731
    return re.sub(r"<(cellXfs|cellStyleXfs)\b.*?</\1>",
                  lambda b: re.sub(r"<xf\b[^>]*>", fix, b.group(0)), st, flags=re.S)


def house(st):
    """styles.xml with the house fonts, borders and fills first and no red. Idempotent."""
    st = st.replace('<fgColor theme="4"/>', f'<fgColor rgb="{NAVY}"/>')
    st = house_first(st, "font", FONTS, "fontId")
    st = house_first(st, "border", BORDERS, "borderId")
    st = house_first(st, "fill", FILLS, "fillId")
    return no_red(st)


def no_red(st):
    """Drop [Red] from every number format, so a record made before 2026-11-18 loses it too."""
    return re.sub(r'(formatCode="[^"]*)', lambda m: m.group(1).replace("[Red]", ""), st)


def numfmt(st, code):
    """Id of a custom number format, adding it if absent."""
    m = re.search(r'<numFmt numFmtId="(\d+)" formatCode="' + re.escape(code.replace('"', "&quot;"))
                  + '"/>', st)
    if m:
        return st, int(m.group(1))
    ids = [int(i) for i in re.findall(r'<numFmt numFmtId="(\d+)"', st)]
    new = max(ids + [180]) + 1
    n = int(re.search(r'<numFmts count="(\d+)"', st).group(1))
    st = st.replace(f'<numFmts count="{n}">', f'<numFmts count="{n + 1}">', 1)
    st = st.replace("</numFmts>", f'<numFmt numFmtId="{new}" formatCode="'
                    f'{code.replace(chr(34), "&quot;")}"/></numFmts>', 1)
    return st, new


def apply(book):
    """Style a loaded Book in place (no save). update_record.py calls this too."""
    # Excel's calculation index lists formula cells; one that no longer exists makes Excel refuse
    # the file. Dropping the index is safe: Excel rebuilds it on open.
    if "xl/calcChain.xml" in book.parts:
        book.drop("xl/calcChain.xml")
        book.sub("xl/_rels/workbook.xml.rels", r'<Relationship [^>]*Target="calcChain\.xml"/>', "")
        book.sub("[Content_Types].xml", r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "")
    st = book.parts["xl/styles.xml"].decode("utf-8")
    st = house(st)
    st, pnl = numfmt(st, PNL)
    st, cost = numfmt(st, COST)
    st, vol = numfmt(st, VOL)
    # the Statement's general number format (177) becomes the P&L one
    st = st.replace(r'formatCode="#,##0.00;\-#,##0.00"', 'formatCode="' + PNL + '"')
    book.parts["xl/styles.xml"] = st.encode("utf-8")

    S = Styles(book)
    x_pnl = S.xf(f'<xf numFmtId="{pnl}" fontId="0" fillId="0" borderId="1" xfId="0" '
                 'applyNumberFormat="1" applyBorder="1"/>')
    x_int = S.xf('<xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" '
                 'applyNumberFormat="1" applyBorder="1"/>')
    x_cost = S.xf(f'<xf numFmtId="{cost}" fontId="0" fillId="0" borderId="1" xfId="0" '
                  'applyNumberFormat="1" applyBorder="1"/>')
    x_vol = S.xf(f'<xf numFmtId="{vol}" fontId="0" fillId="0" borderId="1" xfId="0" '
                 'applyNumberFormat="1" applyBorder="1"/>')
    x_head = S.xf('<xf numFmtId="0" fontId="3" fillId="2" borderId="1" xfId="0" applyFont="1" '
                  'applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" '
                  'vertical="center" wrapText="1"/></xf>')
    x_int_plain = S.xf('<xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" '
                       'applyNumberFormat="1"/>')
    x_bar = str(S.xf(BAR_XML))
    S.save()
    parts = book.sheet_parts()

    # ---------- journal ----------
    j = book.sheet(parts["IMR+PnL"], "IMR+PnL")
    j.expand_shared()
    last = max(j.rows)
    for c, text in zip("BCDEFGHIJKLMNOPQRS", ["Day", "Date", "A/c Val", "Margin", "Daily PnL",
                                          "VIX A/c Val", "VIX Margin", "VIX Daily PnL",
                                          "SR3 TC", "SR1 TC", "ZQ TC", "VX TC", "Rates TC Adj",
                                          "Total TC", "SR3 Vol", "SR1 Vol", "ZQ Vol", "VX Vol"]):
        j.set_text(f"{c}2", text, style=str(x_head))
    for r in range(3, last + 1):
        for c in "DFGI":   # a typed note (G181) keeps its text; the format only affects numbers
            _restyle(j, f"{c}{r}", x_pnl)
        for c in "EH":
            _restyle(j, f"{c}{r}", x_int)
        for c in "JKLMNO":
            _restyle(j, f"{c}{r}", x_cost)
        for c in "PQRS":
            _restyle(j, f"{c}{r}", x_vol)
    j.rows[2][0] = re.sub(r'\s*(?:ht|customHeight)="[^"]*"', "", j.rows[2][0]) + ' ht="30" customHeight="1"'
    j.set_col_widths({2: 6, 3: 10, **{c: 11 for c in range(4, 16)}, **{c: 9 for c in range(16, 20)}})

    # ---------- statement ----------
    s = book.sheet(parts["Statement"], "Statement")
    bars = {4: "C", 13: "C", **{r: "CDEFGHIJKLM" for r in (16, 33, 41)}}  # patch_statement.py's bars
    for r, cols in bars.items():
        for c in cols:
            raw = s.raw(f"{c}{r}")
            if raw and ("inlineStr" in raw or 't="s"' in raw):
                continue                       # a title in the bar keeps its own (right-aligned) style
            s.blank(f"{c}{r}", style=x_bar)
    s.tail = re.sub(r"<conditionalFormatting\b.*?</conditionalFormatting>", "", s.tail, flags=re.S)
    s.tail = re.sub(r'<ext uri="\{78C0D931-6437-407d-A8EE-F0AAD7539E65\}".*?</ext>', "", s.tail, flags=re.S)
    s.tail = re.sub(r"<extLst>\s*</extLst>", "", s.tail)
    s.head = re.sub(r'\s*zoomScale(?:Normal)?="\d+"', "", s.head)

    for sh in (j, s):
        sh.head = re.sub(r'<tabColor rgb="[^"]*"/>', f'<tabColor rgb="{NAVY}"/>', sh.head)



def main(path):
    if path.endswith("PnL Record.xlsx"):
        try:
            open(path, "r+b").close()
        except PermissionError:
            print("PnL Record.xlsx is open in Excel. Close it and run again.")
            return 1
    book = Book(path)
    apply(book)
    if path.endswith("PnL Record.xlsx"):
        from paths import backup
        backup("before restyle")
    tmp = path + ".tmp"
    book.save(tmp)
    os.replace(tmp, path)
    print(f"Styled {path}")
    return 0


def _restyle(sh, ref, xf):
    """Point an existing cell at style xf, or create a styled blank."""
    raw = sh.raw(ref)
    if raw is None:
        sh.blank(ref, style=str(xf))
        return
    new = re.sub(r'\ss="\d+"', "", raw, count=1)
    new = re.sub(r'^<c r="([A-Z]+\d+)"', lambda m: f'<c r="{m.group(1)}" s="{xf}"', new)
    sh._put(ref, new)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "PnL Record.xlsx")))
