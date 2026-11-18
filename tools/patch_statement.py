"""Build the Statement sheet (layout agreed 2026-11-18). Idempotent: owns the whole sheet below
row 1 and can be re-run at any time. Everything is a formula over the journal (IMR+PnL); the one
typed figure is the adjustment (C14), which is kept.

    rows 4-11   A/C SUMMARY: the block copied and sent every day (labels and formats as always;
                nothing sits beside it, so a row copy picks up only this). P&L and account value
                are rates + VIX; margin is the rates account's
    rows 13-14  INPUT: the adjustment
    rows 16-31  MONTHLY, Jan-Dec of the statement's year, then a Year row and a note line
    rows 33-39  QUARTERLY, Q1-Q4, then Year
    rows 41-48  BY YEAR, from the first statement's year
    rows 50-88  CHARTS (his call, 2026-11-18): Total P&L year to date (a line, full width), then
                Total P&L by month and Total TC by month side by side (statement_charts.py)
    row  91     CHART DATA: one row a calendar day of the statement's year (rows 92-458, hidden in
                a group under the bar): the year chart's points, #N/A where there is no statement

Every table has the same columns (his spec, 2026-11-18: total first, for screenshots):
    C Total P&L | D Rates P&L | E VIX P&L | F gap | G Total TC | H Rates TC | I VIX TC | J gap |
    K Total Volume | L Rates Volume | M VIX Volume
P&L is after TC (account value at market is net of costs). Rates P&L = journal F summed over the
window; VIX P&L = the change in the VIX account value (G) over the window, because G has gaps;
Rates TC = journal J+K+L+N, VIX TC = M; Rates Volume = P+Q+R, VIX Volume = S. A VIX cell shows
"-" when no statement in the window carried the VIX account. Months still to come are blank.
"""
import os
import posixpath
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import statement_charts as charts  # noqa: E402
from paths import RECORD, backup, record_is_open  # noqa: E402
from xlsxpatch import Book, Styles  # noqa: E402

J = "'IMR+PnL'!"
LAST = 1000
DATES = f"{J}$C$3:$C${LAST}"
ASOF = "$C$5"
HEADS = {"C": "Total P&L", "D": "Rates P&L", "E": "VIX P&L",
         "G": "Total TC", "H": "Rates TC", "I": "VIX TC",
         "K": "Total Volume", "L": "Rates Volume", "M": "VIX Volume"}
GAPS = "FJ"
SPAN = "CDEFGHIJKLM"  # every column a table's bar covers, after the label in B
VOLUME = "KLM"
SMALL = "DEHILM"  # the Rates and VIX columns: a smaller font than Total's
SMALL_PT = 9      # the house size is 11
WIDTHS = {1: 2.4, 2: 16, 3: 13, 4: 13, 5: 13, 6: 1.5, 7: 12, 8: 12, 9: 12, 10: 1.5,
          11: 13, 12: 13, 13: 13, 14: 2.4}
CHART_ROWS = ((49, 68), (69, 88))   # 0-based anchor rows: the year chart, then the two monthly ones
DATA_BAR, DATA_HEAD = 91, 92
DATA_FIRST, DATA_DAYS = 93, 366
DATA_LAST = DATA_FIRST + DATA_DAYS - 1
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')


def col(c):
    return f"{J}${c}$3:${c}${LAST}"


def sif(c, s, e):
    return f'SUMIFS({col(c)},{DATES},">="&{s},{DATES},"<="&{e})'


def last_of(c):
    """The journal's last value in column c (on the last row that has an account value)."""
    return f'IFERROR(_xlfn.XLOOKUP(TRUE,{col("D")}<>"",{col(c)},"",0,-1),"")'


def vix_level(x):
    """VIX account value on or before date x; 0 before the first one."""
    return f"_xlfn.XLOOKUP(1,ISNUMBER({col('G')})*({DATES}<={x}),{col('G')},0,0,-1)"


def vix_days(s, e):
    return f"SUMPRODUCT(({DATES}>={s})*({DATES}<={e})*ISNUMBER({col('G')}))"


def row_formulas(r, s, e, future):
    """Column -> formula for window s..e (e capped at the statement date)."""
    e = f"MIN({e},{ASOF})"
    blank = f'OR({ASOF}="",{future})'
    vix_none = f"{vix_days(s, e)}=0"
    f = {
        "D": sif("F", s, e),
        "E": f'IF({vix_none},"-",{vix_level(e)}-{vix_level(f"({s}-1)")})',
        "C": f"D{r}+N(E{r})",
        "H": "+".join(sif(c, s, e) for c in "JKLN"),
        "I": f'IF({vix_none},"-",{sif("M", s, e)})',
        "G": f"H{r}+N(I{r})",
        "L": "+".join(sif(c, s, e) for c in "PQR"),
        "M": f'IF({vix_none},"-",{sif("S", s, e)})',
        "K": f"L{r}+N(M{r})",
    }
    return {c: f'IF({blank},"",{v})' for c, v in f.items()}


# ---------- charts: this script owns the Statement's drawing and rebuilds it on every run ----------

def _rels_of(part):
    d, f = part.rsplit("/", 1)
    return f"{d}/_rels/{f}.rels"


def _resolve(base, target):
    return target.lstrip("/") if target.startswith("/") else posixpath.normpath(
        posixpath.join(posixpath.dirname(base), target))


def _put(book, part, text):
    if part not in book.parts:
        book.order.append(part)
    book.parts[part] = text.encode("utf-8")


def _drop(book, part, ct):
    """Drop a part, everything its rels point to inside the package, and their content types.
    Excel renames the drawing and charts and adds style and colour parts when it saves, so the
    old drawing is found through the sheet's rels, never by name."""
    rels = _rels_of(part)
    if rels in book.parts:
        for rel in re.findall(r"<Relationship\b[^>]*>", book.parts[rels].decode("utf-8")):
            if 'TargetMode="External"' not in rel:
                _drop(book, _resolve(part, re.search(r'Target="([^"]+)"', rel).group(1)), ct)
        book.drop(rels)
    book.drop(part)
    ct[0] = re.sub(rf'<Override PartName="/{re.escape(part)}"[^>]*/>', "", ct[0])


def _free(book, fmt):
    n = 1
    while fmt.format(n) in book.parts:
        n += 1
    return fmt.format(n)


def _px(width):
    """Pixel width of a column of the given character width (Excel's rule, 7 px a digit)."""
    return int((256 * width + 18) / 256 * 7)


def _at(x):
    """(0-based column, offset in EMU) of a point x px right of column B's left edge."""
    for c in range(2, 14):
        w = _px(WIDTHS[c])
        if x < w:
            return c - 1, int(x * 9525)
        x -= w
    return 13, 0


def put_charts(book, part, sh, specs):
    """Replace the sheet's drawing with charts: specs = [(chart xml, (col, off), (col, off), (row0, row1))]."""
    rels_p = _rels_of(part)
    rels = book.parts[rels_p].decode("utf-8") if rels_p in book.parts else PKG_RELS + "</Relationships>"
    ct = [book.parts["[Content_Types].xml"].decode("utf-8")]
    m = re.search(r'<drawing r:id="([^"]+)"\s*/>', sh.tail)
    if m:
        rel = re.search(rf'<Relationship\b[^>]*\bId="{m.group(1)}"[^>]*>', rels)
        if rel:
            _drop(book, _resolve(part, re.search(r'Target="([^"]+)"', rel.group(0)).group(1)), ct)
            rels = rels.replace(rel.group(0), "")
        sh.tail = sh.tail.replace(m.group(0), "")
    drawing = _free(book, "xl/drawings/drawing{}.xml")
    anchors, drels = [], []
    for i, (xml, (c0, o0), (c1, o1), (r0, r1)) in enumerate(specs):
        chart = _free(book, "xl/charts/chart{}.xml")
        _put(book, chart, xml)
        ct[0] = ct[0].replace("</Types>", f'<Override PartName="/{chart}" ContentType="application/'
                                          'vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>')
        drels.append(f'<Relationship Id="rId{i + 1}" Type="{REL}/chart" Target="../charts/{chart.rsplit("/", 1)[1]}"/>')
        anchors.append(
            f'<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>{c0}</xdr:col><xdr:colOff>{o0}</xdr:colOff>'
            f'<xdr:row>{r0}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>{c1}</xdr:col>'
            f'<xdr:colOff>{o1}</xdr:colOff><xdr:row>{r1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
            f'<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{i + 2}" name="Chart {i + 1}"/>'
            '<xdr:cNvGraphicFramePr><a:graphicFrameLocks/></xdr:cNvGraphicFramePr></xdr:nvGraphicFramePr>'
            '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic>'
            '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
            f'<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="{REL}" '
            f'r:id="rId{i + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/>'
            '</xdr:twoCellAnchor>')
    _put(book, drawing, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<xdr:wsDr xmlns:xdr='
                        '"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a='
                        '"http://schemas.openxmlformats.org/drawingml/2006/main">' + "".join(anchors) + "</xdr:wsDr>")
    _put(book, _rels_of(drawing), PKG_RELS + "".join(drels) + "</Relationships>")
    ct[0] = ct[0].replace("</Types>", f'<Override PartName="/{drawing}" ContentType="application/'
                                      'vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
    rid = "rId" + str(max([int(i) for i in re.findall(r'Id="rId(\d+)"', rels)] + [0]) + 1)
    rels = rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="{REL}/drawing" '
                                            f'Target="../drawings/{drawing.rsplit("/", 1)[1]}"/></Relationships>')
    _put(book, rels_p, rels)
    book.parts["[Content_Types].xml"] = ct[0].encode("utf-8")
    # the worksheet's <drawing> goes after pageSetup and before these, in schema order
    nxt = re.search(r"<(?:legacyDrawing|legacyDrawingHF|drawingHF|picture|oleObjects|controls|"
                    r"webPublishItems|tableParts|extLst)\b|</worksheet>", sh.tail)
    sh.tail = sh.tail[:nxt.start()] + f'<drawing r:id="{rid}"/>' + sh.tail[nxt.start():]


def apply(book):
    part = book.sheet_parts()["Statement"]
    sh = book.sheet(part, "Statement")
    # P&L and TC shown to the unit, never red (his call, 2026-11-18); the cell keeps the exact figure
    import patch_style
    # house fonts, borders and fills first, so the ids below mean what they say even after Excel saved
    styles_xml, whole = patch_style.numfmt(patch_style.house(book.parts["xl/styles.xml"].decode("utf-8")),
                                           r'#,##0;\-#,##0')
    book.parts["xl/styles.xml"] = styles_xml.encode("utf-8")
    st = Styles(book)
    x = lambda xml: str(st.xf(xml))  # noqa: E731
    # Rates and VIX columns in a smaller font, Total at full size (his call, 2026-11-18)
    small = {False: st.font(f'<font><sz val="{SMALL_PT}"/><color theme="1"/>{patch_style.F}</font>'),
             True: st.font(f'<font><b/><sz val="{SMALL_PT}"/><color theme="1"/>{patch_style.F}</font>')}

    def pair(fmt, bold, border, align="right"):
        """(full-size style, small style) for one kind of cell."""
        tail = (f'fillId="0" borderId="{border}" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" '
                f'applyAlignment="1"><alignment horizontal="{align}"/></xf>')
        return (x(f'<xf numFmtId="{fmt}" fontId="{1 if bold else 0}" {tail}'),
                x(f'<xf numFmtId="{fmt}" fontId="{small[bold]}" {tail}'))

    num, vol = pair(whole, False, 0), pair(3, False, 0)
    tot_num, tot_vol = pair(whole, True, 5), pair(3, True, 5)
    head = pair(0, True, 4)
    tot_lab = x('<xf numFmtId="0" fontId="1" fillId="0" borderId="5" xfId="0" applyFont="1" applyBorder="1"/>')
    note = x('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">'
             '<alignment horizontal="left"/></xf>')
    BAR = x(patch_style.BAR_XML)
    SUB_L = x('<xf numFmtId="0" fontId="1" fillId="0" borderId="4" xfId="0" applyFont="1" applyBorder="1"/>')
    MONTH = x('<xf numFmtId="17" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" '
              'applyAlignment="1"><alignment horizontal="left"/></xf>')
    DAY = x('<xf numFmtId="15" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" '
            'applyAlignment="1"><alignment horizontal="left"/></xf>')
    st.save()

    # keep the send block's own cells (labels, styles, formats) and the adjustment; rewrite the
    # figures so the block stands on its own (it used to read a "statement values" block)
    keep = {f"B{r}" for r in range(2, 12)} | {f"C{r}" for r in range(5, 12)} | {"B14", "C14"}
    for r in list(sh.rows):
        for c in list(sh.rows[r][1]):
            ref = f"{chr(64 + c)}{r}" if c <= 26 else None
            if ref not in keep:
                del sh.rows[r][1][c]
    # P&L lines are rates + VIX (his call, 2026-11-18); margin stays the rates account's alone.
    # VIX is read as in the tables below: its account value on or before a date, 0 before the first.
    q_start = f"DATE(YEAR({ASOF}),FLOOR(MONTH({ASOF})-1,3)+1,1)"
    sh.set_text("B4", "A/C SUMMARY")
    send = {
        "C5": last_of("C"),
        "C6": f'IF({ASOF}="","",{sif("F", q_start, ASOF)}+{vix_level(ASOF)}-{vix_level(f"({q_start}-1)")})',
        "C7": f'IF({ASOF}="","",N({last_of("F")})+N({last_of("I")}))',
        "C8": f'IF({ASOF}="","",N({last_of("D")})+{vix_level(ASOF)})',
        "C9": last_of("E"),
        "C10": "$C$14",
        "C11": "C8+C10",
    }
    for ref, f in send.items():
        sh.set_formula(ref, f)  # keeps the cell's style, so the sent formats are unchanged

    def bar(r, title, span=SPAN):
        sh.set_text(f"B{r}", title, style=BAR)
        for c in span:
            sh.blank(f"{c}{r}", style=BAR)

    def heads(r, first):
        sh.set_text(f"B{r}", first, style=SUB_L)
        for c, h in HEADS.items():
            sh.set_text(f"{c}{r}", h, style=head[c in SMALL])

    def figures(r, s, e, future):
        for c, f in row_formulas(r, s, e, future).items():
            sh.set_formula(f"{c}{r}", f, style=(vol if c in VOLUME else num)[c in SMALL])

    def total(r, first, last, label="Year"):
        sh.set_text(f"B{r}", label, style=tot_lab)
        for c in HEADS:
            sh.set_formula(f"{c}{r}", f"SUM({c}{first}:{c}{last})",
                           style=(tot_vol if c in VOLUME else tot_num)[c in SMALL])

    sh.blank("C4", style=BAR)  # B4 keeps its "A/C SUMMARY" title; the bar spans the block only
    bar(13, "INPUT", span="C")

    # monthly
    bar(16, "MONTHLY")
    heads(17, "Month")
    for k in range(12):
        r = 18 + k
        sh.set_formula(f"B{r}", f'IF({ASOF}="","",DATE(YEAR({ASOF}),{k + 1},1))', style=MONTH)
        figures(r, f"$B{r}", f"EOMONTH($B{r},0)", f"$B{r}>{ASOF}")
    total(30, 18, 29)
    first_vix = f"_xlfn.XLOOKUP(TRUE,ISNUMBER({col('G')}),{DATES},\"\",0,1)"
    sh.set_formula("B31", f'IFERROR(IF(AND({first_vix}<>"",{first_vix}>DATE(YEAR({first_vix}),1,10)),'
                          f'"VIX: no daily statements before "&TEXT({first_vix},"d mmm yyyy")&'
                          f'"; that month\'s VIX P&L and TC include everything before it, its VIX volume does not.",""),"")',
                   style=note)

    # quarterly
    bar(33, "QUARTERLY")
    heads(34, "Quarter")
    for k in range(4):
        r, qs = 35 + k, f"DATE(YEAR({ASOF}),{3 * k + 1},1)"
        sh.set_formula(f"B{r}", f'IF({ASOF}="","","Q{k + 1} "&YEAR({ASOF}))')
        figures(r, qs, f"EOMONTH({qs},2)", f"{qs}>{ASOF}")
    total(39, 35, 38)

    # by year
    bar(41, "BY YEAR")
    heads(42, "Year")
    y0 = f"YEAR(MIN({DATES}))"
    for k in range(6):
        r = 43 + k
        yr = f"({y0}+{k})"
        sh.set_formula(f"B{r}", f'IF(OR({ASOF}="",{yr}>YEAR({ASOF})),"",{yr})', style=note)
        figures(r, f"DATE({yr},1,1)", f"DATE({yr},12,31)", f"{yr}>YEAR({ASOF})")

    # chart data: Total P&L year to date for each calendar day of the statement's year, by the
    # tables' own rule (rates: journal F summed; VIX: the change in its account value); #N/A on a
    # day with no statement or after the statement date, which the line joins over or stops at
    bar(DATA_BAR, "CHART DATA", span="C")
    sh.set_text(f"B{DATA_HEAD}", "Day", style=SUB_L)
    sh.set_text(f"C{DATA_HEAD}", "Total P&L YTD", style=head[False])
    jan1 = f"DATE(YEAR({ASOF}),1,1)"
    for k in range(DATA_DAYS):
        r = DATA_FIRST + k
        # capped at 31 Dec: in a 365-day year the last row repeats it, so the axis ends with the year
        sh.set_formula(f"B{r}", f'IF({ASOF}="","",MIN({jan1}+{k},DATE(YEAR({ASOF}),12,31)))', style=DAY)
        sh.set_formula(f"C{r}", f'IF(OR({ASOF}="",B{r}>{ASOF},COUNTIFS({DATES},B{r},{col("D")},"<>")=0),NA(),'
                                f'{sif("F", jan1, f"B{r}")}+{vix_level(f"B{r}")}-{vix_level(f"({jan1}-1)")})',
                       style=num[False])
    for r in sh.rows:  # the data rows sit in a collapsed group under the bar; nothing else is hidden
        if r >= 50:
            sh.rows[r][0] = f' r="{r}"' + (' hidden="1" outlineLevel="1"' if DATA_HEAD <= r <= DATA_LAST
                                           else ' collapsed="1"' if r == DATA_BAR else "")
    sh.head = re.sub(r'\s*outlineLevelRow="\d+"', "", sh.head)
    sh.head = re.sub(r"<sheetFormatPr\b", '<sheetFormatPr outlineLevelRow="1"', sh.head, count=1)
    sh.head = re.sub(r"<outlinePr\b[^>]*/>", "", sh.head)
    if re.search(r"<sheetPr\b[^>]*/>", sh.head):
        sh.head = re.sub(r"<sheetPr\b([^>]*)/>", r"<sheetPr\1></sheetPr>", sh.head, count=1)
    if "<sheetPr" not in sh.head:
        sh.head = re.sub(r"(<worksheet\b[^>]*>)", r"\1<sheetPr></sheetPr>", sh.head, count=1)
    after_tab = re.search(r"<sheetPr\b[^>]*>(?:\s*<tabColor\b[^>]*/>)?", sh.head)
    sh.head = sh.head[:after_tab.end()] + '<outlinePr summaryBelow="0"/>' + sh.head[after_tab.end():]

    # the charts: the year line full width, then the two monthly ones side by side
    whole = sum(_px(WIDTHS[c]) for c in range(2, 14))
    mid, gap = whole / 2, 6
    (y0, y1), (m0, m1) = CHART_ROWS
    put_charts(book, part, sh, [
        (charts.year_pnl(f"Statement!$B${DATA_FIRST}:$B${DATA_LAST}", f"Statement!$C${DATA_FIRST}:$C${DATA_LAST}"),
         (1, 0), (13, 0), (y0, y1)),
        (charts.monthly("Total P&amp;L by month", "Total P&amp;L", "Statement!$B$18:$B$29",
                        "Statement!$C$18:$C$29", charts.NAVY, 620000001),
         (1, 0), _at(mid - gap), (m0, m1)),
        (charts.monthly("Total TC by month", "Total TC", "Statement!$B$18:$B$29",
                        "Statement!$G$18:$G$29", charts.GREY, 630000001),
         _at(mid + gap), (13, 0), (m0, m1)),
    ])

    sh.set_col_widths(WIDTHS)
    sh.head = re.sub(r'<selection [^>]*/>', '<selection activeCell="C14" sqref="C14"/>', sh.head)


def main():
    if record_is_open():
        print("PnL Record.xlsx is open in Excel. Close it and run again.")
        return 1
    book = Book(RECORD)
    apply(book)
    import patch_style
    patch_style.apply(book)
    book.sub("xl/workbook.xml", r"<calcPr ((?:(?!fullCalcOnLoad)[^/])*?)/>",
             r'<calcPr \1 fullCalcOnLoad="1"/>', count=1)
    backup("before statement rebuild")
    book.save(RECORD + ".tmp")
    os.replace(RECORD + ".tmp", RECORD)
    print("Statement built.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
