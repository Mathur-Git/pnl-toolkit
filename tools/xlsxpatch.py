"""Minimal, dependency-light OOXML worksheet patcher.

Edits cells inside an .xlsx package without going through openpyxl's
read/write cycle, so charts, form controls, VML, comments, web extensions
and the Python-in-Excel part all survive byte-identical.
"""
import re, html, zipfile, shutil
from openpyxl.utils import get_column_letter as gcl, column_index_from_string as cidx
from openpyxl.formula.translate import Translator

CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
ROW_RE = re.compile(r'<row([^>]*?)(?:/>|>(.*?)</row>)', re.S)


def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


class Sheet:
    """Parsed sheetData: {row:int -> (row_attrs, {col:int -> cellxml})}."""

    def __init__(self, xml, name=''):
        self.name = name
        self.xml = xml
        m = re.search(r'(<sheetData\s*/>|<sheetData>.*?</sheetData>)', xml, re.S)
        self.head = xml[:m.start()]
        self.tail = xml[m.end():]
        body = m.group(1)
        self.rows = {}
        if not body.startswith('<sheetData/>'):
            inner = body[len('<sheetData>'):-len('</sheetData>')]
            for rm in ROW_RE.finditer(inner):
                attrs, content = rm.group(1), rm.group(2) or ''
                rnum = int(re.search(r'r="(\d+)"', attrs).group(1))
                attrs = re.sub(r'\s*spans="[^"]*"', '', attrs)
                cells = {}
                for cm in CELL_RE.finditer(content):
                    cells[cidx(cm.group(1))] = cm.group(0)
                self.rows[rnum] = [attrs, cells]

    # ---------- helpers ----------
    def _cell(self, ref):
        m = re.match(r'([A-Z]+)(\d+)$', ref)
        return cidx(m.group(1)), int(m.group(2))

    def raw(self, ref):
        c, r = self._cell(ref)
        return self.rows.get(r, [None, {}])[1].get(c)

    def style_of(self, ref):
        raw = self.raw(ref)
        if not raw:
            return None
        m = re.search(r'\ss="(\d+)"', raw)
        return m.group(1) if m else None

    def _put(self, ref, xml):
        c, r = self._cell(ref)
        if r not in self.rows:
            self.rows[r] = [f' r="{r}"', {}]
        self.rows[r][1][c] = xml

    # ---------- shared-formula expansion ----------
    def expand_shared(self):
        masters, fails = {}, []
        for r in sorted(self.rows):
            for c in sorted(self.rows[r][1]):
                raw = self.rows[r][1][c]
                fm = re.search(r'<f([^>]*?)>(.*?)</f>', raw, re.S)
                if not fm:
                    continue
                a = fm.group(1)
                if 't="shared"' in a and 'ref=' in a:
                    si = re.search(r'si="(\d+)"', a).group(1)
                    masters[si] = (f'{gcl(c)}{r}', html.unescape(fm.group(2)))
        for r in sorted(self.rows):
            for c in sorted(self.rows[r][1]):
                raw = self.rows[r][1][c]
                fm = re.search(r'<f([^>]*?)(?:/>|>(.*?)</f>)', raw, re.S)
                if not fm or 't="shared"' not in fm.group(1):
                    continue
                si = re.search(r'si="(\d+)"', fm.group(1)).group(1)
                if si not in masters:
                    fails.append(f'{self.name}!{gcl(c)}{r} orphan si={si}')
                    continue
                origin, ftext = masters[si]
                ref = f'{gcl(c)}{r}'
                try:
                    newf = (ftext if ref == origin
                            else Translator('=' + ftext, origin=origin)
                            .translate_formula(ref)[1:])
                except Exception as e:                       # pragma: no cover
                    fails.append(f'{self.name}!{ref} {e}')
                    continue
                keep = ''
                if 'ca="1"' in fm.group(1):
                    keep = ' ca="1"'
                self.rows[r][1][c] = (raw[:fm.start()] +
                                      f'<f{keep}>{esc(newf)}</f>' + raw[fm.end():])
        return fails

    # ---------- edits ----------
    def set_formula(self, ref, formula, style=None, keep_f_attrs=False, f_attrs=''):
        """Write a formula; drop any cached value and value-type."""
        raw = self.raw(ref)
        fattrs = f_attrs
        if raw and keep_f_attrs:
            fm = re.search(r'<f([^>]*?)(?:/>|>)', raw, re.S)
            if fm:
                fattrs = fm.group(1)
        s = style if style is not None else (self.style_of(ref) or None)
        sat = f' s="{s}"' if s is not None else ''
        self._put(ref, f'<c r="{ref}"{sat}><f{fattrs}>{esc(formula)}</f></c>')

    def set_number(self, ref, value, style=None):
        s = style if style is not None else (self.style_of(ref) or None)
        sat = f' s="{s}"' if s is not None else ''
        self._put(ref, f'<c r="{ref}"{sat}><v>{value}</v></c>')

    def set_text(self, ref, text, style=None):
        s = style if style is not None else (self.style_of(ref) or None)
        sat = f' s="{s}"' if s is not None else ''
        self._put(ref, f'<c r="{ref}"{sat} t="inlineStr"><is><t xml:space="preserve">'
                       f'{esc(text)}</t></is></c>')

    def blank(self, ref, style=None):
        """Keep the cell (and its formatting) but remove content."""
        s = style if style is not None else self.style_of(ref)
        sat = f' s="{s}"' if s is not None else ''
        self._put(ref, f'<c r="{ref}"{sat}/>')

    def sub_formula(self, ref, pattern, repl, count=0):
        """Regex-replace inside an existing cell's formula, keeping attributes."""
        raw = self.raw(ref)
        if not raw:
            return False
        fm = re.search(r'<f([^>]*?)>(.*?)</f>', raw, re.S)
        if not fm:
            return False
        old = html.unescape(fm.group(2))
        new = re.sub(pattern, repl, old, count=count)
        if new == old:
            return False
        cell = raw[:fm.start()] + f'<f{fm.group(1)}>{esc(new)}</f>' + raw[fm.end():]
        cell = re.sub(r'<v>.*?</v>', '', cell, flags=re.S)
        cell = re.sub(r'(<c r="[A-Z]+\d+"[^>]*?)\st="(?:str|e|b|s)"', r'\1', cell)
        self._put(ref, cell)
        return True

    def set_col_widths(self, widths):
        """widths: {column_index: width}. Splits/merges existing <col> runs as needed."""
        m = re.search(r'<cols>(.*?)</cols>', self.head, re.S)
        entries = []
        if m:
            for cm in re.finditer(r'<col ([^/]*?)/>', m.group(1)):
                a = cm.group(1)
                lo = int(re.search(r'min="(\d+)"', a).group(1))
                hi = int(re.search(r'max="(\d+)"', a).group(1))
                rest = re.sub(r'\s*(?:min|max|width|customWidth)="[^"]*"', '', a).strip()
                w = re.search(r'width="([\d.]+)"', a)
                entries.append([lo, hi, float(w.group(1)) if w else None, rest])
        out = []
        for lo, hi, w, rest in entries:                       # split around overrides
            cur = lo
            while cur <= hi:
                if cur in widths:
                    out.append([cur, cur, widths[cur], rest]); cur += 1
                else:
                    nxt = cur
                    while nxt <= hi and nxt not in widths:
                        nxt += 1
                    out.append([cur, nxt - 1, w, rest]); cur = nxt
        covered = {i for lo, hi, _, _ in out for i in range(lo, hi + 1)}
        for c, w in sorted(widths.items()):                   # columns with no prior entry
            if c not in covered:
                out.append([c, c, w, ''])
        out.sort(key=lambda e: e[0])
        merged = []
        for e in out:                                          # merge adjacent identical runs
            if merged and merged[-1][1] + 1 == e[0] and merged[-1][2:] == e[2:]:
                merged[-1][1] = e[1]
            else:
                merged.append(e)
        xml = ''.join(
            f'<col min="{lo}" max="{hi}"'
            + (f' width="{w:g}" customWidth="1"' if w is not None else '')
            + (f' {rest}' if rest else '') + '/>'
            for lo, hi, w, rest in merged)
        if m:
            self.head = self.head.replace(m.group(0), f'<cols>{xml}</cols>')
        else:
            self.head = re.sub(r'(<sheetData)', f'<cols>{xml}</cols>\\1', self.head, count=1)

    # ---------- serialise ----------
    def tostring(self):
        out, mn, mx, rmin, rmax = [], None, None, None, None
        for r in sorted(self.rows):
            attrs, cells = self.rows[r]
            if not cells:
                out.append(f'<row{attrs}/>')
                continue
            cols = sorted(cells)
            out.append(f'<row{attrs} spans="{cols[0]}:{cols[-1]}">'
                       + ''.join(cells[c] for c in cols) + '</row>')
            mn = cols[0] if mn is None else min(mn, cols[0])
            mx = cols[-1] if mx is None else max(mx, cols[-1])
            rmin = r if rmin is None else min(rmin, r)
            rmax = r if rmax is None else max(rmax, r)
        head = self.head
        if mn is not None:
            dim = f'{gcl(mn)}{rmin}:{gcl(mx)}{rmax}'
            head = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="{dim}"/>', head)
        return head + '<sheetData>' + ''.join(out) + '</sheetData>' + self.tail


class Styles:
    """Append-only editor for fonts / cellXfs in styles.xml, idempotent by identity.

    Adding the same font or xf twice returns the index of the one already there, so
    an in-place patch script can be re-run without growing styles.xml every time.
    """

    def __init__(self, book):
        self.bk = book
        self.st = book.parts['xl/styles.xml'].decode('utf-8')

    def _add(self, tag, item_re, xml):
        s, e = self.st.index(f'<{tag} '), self.st.index(f'</{tag}>')
        blk = self.st[s:e]
        at = blk.find(xml)
        if at >= 0:                                     # already present -> reuse
            return len(re.findall(item_re, blk[:at]))
        n = int(re.search(rf'<{tag} count="(\d+)"', blk).group(1))
        assert n == len(re.findall(item_re, blk)), f'{tag} count out of step'
        self.st = (self.st[:s]
                   + blk.replace(f'<{tag} count="{n}"', f'<{tag} count="{n + 1}"', 1)
                   + xml + self.st[e:])
        return n

    def font(self, xml):
        return self._add('fonts', r'<font[ />]', xml)

    def xf(self, xml):
        return self._add('cellXfs', r'<xf[ />]', xml)

    def xf_xml(self, idx):
        """The raw <xf> at cellXfs index idx (handles the nested extLst form)."""
        blk = self.st[self.st.index('<cellXfs '):self.st.index('</cellXfs>')]
        blk = blk[blk.index('>') + 1:]
        out, i = [], 0
        while True:
            m = re.compile(r'<xf[ />]').search(blk, i)
            if not m:
                break
            sc = re.match(r'<xf[^>]*/>', blk[m.start():])
            end = m.start() + sc.end() if sc else blk.index('</xf>', m.start()) + 5
            out.append(blk[m.start():end])
            i = end
        return out[idx]

    def save(self):
        self.bk.parts['xl/styles.xml'] = self.st.encode('utf-8')


class Book:
    def __init__(self, path):
        self.parts = {}
        self.order = []
        with zipfile.ZipFile(path) as z:
            for n in z.namelist():
                self.parts[n] = z.read(n)
                self.order.append(n)
        self.sheets = {}

    def sheet(self, part, name=''):
        if part not in self.sheets:
            self.sheets[part] = Sheet(self.parts[part].decode('utf-8'), name)
        return self.sheets[part]

    def sheet_parts(self):
        """{sheet name: worksheet part}, read off this package.

        Never hard-code sheetN.xml: Excel renumbers the parts when it re-saves,
        and v1 and v2 no longer agree (v1 sheet4 = ExpSheet, v2 sheet4 = Vix).
        A script that reads one book and writes another must resolve each.
        """
        wb = self.parts['xl/workbook.xml'].decode('utf-8')
        rels = self.parts['xl/_rels/workbook.xml.rels'].decode('utf-8')
        rid = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
        return {m.group(1): 'xl/' + rid[m.group(2)].lstrip('/').replace('xl/', '', 1)
                for m in re.finditer(r'<sheet name="([^"]+)"[^>]*r:id="([^"]+)"', wb)}

    def drop(self, part):
        if part in self.parts:
            del self.parts[part]
            self.order.remove(part)

    def sub(self, part, pattern, repl, count=0):
        d = self.parts[part].decode('utf-8')
        n = re.sub(pattern, repl, d, count=count)
        self.parts[part] = n.encode('utf-8')
        return n != d

    def save(self, path):
        for part, sh in self.sheets.items():
            self.parts[part] = sh.tostring().encode('utf-8')
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
            for n in self.order:
                z.writestr(n, self.parts[n])
