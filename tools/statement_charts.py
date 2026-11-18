"""Chart XML for the Statement sheet's three charts (his call, 2026-11-18), placed by
patch_statement.py. Adapted from ..\\excel\\tools\\statement_charts.py.

Lean on purpose: no cached values and no chart style or colour parts; the record is saved with
fullCalcOnLoad="1", so Excel fills the charts on open. His look: horizontal and vertical gridlines,
an outline round each chart so it stands apart, never red. One series a chart, so no legend: the
title names it.
"""

NS = ('xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
NAVY = "1F4E79"
GREY = "808080"
MONEY = '#,##0;\\-#,##0'
TEXT = '<a:solidFill><a:srgbClr val="404040"/></a:solidFill>'


def _ln(rgb, w=9525):
    return f'<a:ln w="{w}" cap="flat"><a:solidFill><a:srgbClr val="{rgb}"/></a:solidFill><a:round/></a:ln>'


def _grid(tag, rgb):
    return f'<c:{tag}><c:spPr>{_ln(rgb)}</c:spPr></c:{tag}>'


def _title(text):
    return ('<c:title><c:tx><c:rich><a:bodyPr rot="0" vert="horz" anchor="ctr" anchorCtr="1"/><a:lstStyle/>'
            f'<a:p><a:pPr><a:defRPr sz="1100" b="1">{TEXT}<a:latin typeface="+mn-lt"/></a:defRPr></a:pPr>'
            f'<a:r><a:rPr lang="en-GB" sz="1100" b="1">{TEXT}</a:rPr><a:t>{text}</a:t></a:r></a:p></c:rich></c:tx>'
            '<c:overlay val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:title>'
            '<c:autoTitleDeleted val="0"/>')


def _txpr(sz=900, rot="-60000000"):
    return (f'<c:txPr><a:bodyPr rot="{rot}" vert="horz"/><a:lstStyle/><a:p><a:pPr>'
            f'<a:defRPr sz="{sz}">{TEXT}<a:latin typeface="+mn-lt"/></a:defRPr></a:pPr>'
            '<a:endParaRPr lang="en-GB"/></a:p></c:txPr>')


def _axis_common(fmt):
    return (f'<c:numFmt formatCode="{fmt}" sourceLinked="0"/><c:majorTickMark val="none"/>'
            '<c:minorTickMark val="none"/><c:tickLblPos val="low"/>'
            f'<c:spPr><a:noFill/>{_ln("BFBFBF")}</c:spPr>{_txpr()}')


def _val_ax(axid, cross, between):
    return (f'<c:valAx><c:axId val="{axid}"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
            f'<c:delete val="0"/><c:axPos val="l"/>{_grid("majorGridlines", "D9D9D9")}'
            f'{_axis_common(MONEY)}<c:crossAx val="{cross}"/><c:crosses val="autoZero"/>'
            f'<c:crossBetween val="{between}"/></c:valAx>')


def _space(inner):
    """The chart frame: white, with a navy outline so it stands apart from the sheet."""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            f'<c:chartSpace {NS}><c:date1904 val="0"/><c:lang val="en-GB"/>'
            f'<c:roundedCorners val="0"/><c:chart>{inner}</c:chart>'
            f'<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>{_ln(NAVY, 15875)}</c:spPr>'
            '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:latin typeface="+mn-lt"/>'
            '</a:defRPr></a:pPr><a:endParaRPr lang="en-GB"/></a:p></c:txPr></c:chartSpace>')


def year_pnl(cat, val):
    """Line: Total P&L, year to date, one point per statement day. The axis runs 1 Jan to 31 Dec
    with a darker vertical line at each quarter's start and a light one at each month's.
    Days without a statement are #N/A, which a line joins over; the chart data rows are hidden,
    so plotVisOnly is off."""
    ser = ('<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Total P&amp;L</c:v></c:tx>'
           f'<c:spPr>{_ln(NAVY, 22225)}<a:effectLst/></c:spPr><c:marker><c:symbol val="none"/></c:marker>'
           f'<c:cat><c:numRef><c:f>{cat}</c:f></c:numRef></c:cat>'
           f'<c:val><c:numRef><c:f>{val}</c:f></c:numRef></c:val><c:smooth val="0"/></c:ser>')
    date_ax = ('<c:dateAx><c:axId val="610000001"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
               '<c:delete val="0"/><c:axPos val="b"/>'
               f'{_grid("majorGridlines", "A6A6A6")}{_grid("minorGridlines", "E7E6E6")}'
               f'{_axis_common("mmm")}<c:crossAx val="610000002"/><c:crosses val="autoZero"/>'
               '<c:auto val="1"/><c:lblOffset val="100"/><c:baseTimeUnit val="days"/>'
               '<c:majorUnit val="3"/><c:majorTimeUnit val="months"/>'
               '<c:minorUnit val="1"/><c:minorTimeUnit val="months"/></c:dateAx>')
    return _space(
        _title('Total P&amp;L, year to date')
        + '<c:plotArea><c:layout/><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
        + ser + '<c:marker val="1"/><c:axId val="610000001"/><c:axId val="610000002"/></c:lineChart>'
        + date_ax + _val_ax("610000002", "610000001", "midCat")
        + '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>'
          '<c:plotVisOnly val="0"/><c:dispBlanksAs val="span"/>')


KILO = '#,##0.0,"k";\\-#,##0.0,"k";;'  # 5,083.56 shows 5.1k (his call, 2026-11-18); blank months show nothing


def monthly(title, name, cat, val, rgb, axid):
    """Columns, one a month, Jan to Dec, each labelled in thousands; months to come are blank and
    draw nothing."""
    labels = (f'<c:dLbls><c:numFmt formatCode="{KILO.replace(chr(34), "&quot;")}" sourceLinked="0"/>'
              f'<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>{_txpr(800, rot="0")}<c:dLblPos val="outEnd"/>'
              '<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/>'
              '<c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')
    ser = (f'<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>{name}</c:v></c:tx>'
           f'<c:spPr><a:solidFill><a:srgbClr val="{rgb}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>'
           f'<c:invertIfNegative val="0"/>{labels}'
           f'<c:cat><c:numRef><c:f>{cat}</c:f></c:numRef></c:cat>'
           f'<c:val><c:numRef><c:f>{val}</c:f></c:numRef></c:val></c:ser>')
    a, b = str(axid), str(axid + 1)
    cat_ax = (f'<c:catAx><c:axId val="{a}"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
              f'<c:delete val="0"/><c:axPos val="b"/>{_grid("majorGridlines", "D9D9D9")}'
              f'{_axis_common("mmm")}<c:crossAx val="{b}"/><c:crosses val="autoZero"/>'
              '<c:auto val="0"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>')
    return _space(
        _title(title)
        + '<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>'
          '<c:varyColors val="0"/>' + ser
        + f'<c:gapWidth val="60"/><c:axId val="{a}"/><c:axId val="{b}"/></c:barChart>'
        + cat_ax + _val_ax(b, a, "between")
        + '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>'
          '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>')
