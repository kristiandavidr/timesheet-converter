"""Fill the FORMAT Timesheet template from parsed Jibble data.

Template anchors (see make_clean_template.py):
  MASTER DATA     rows 1-6 header, rows 7-11 = prototype person block (styles)
  01 REKAP ABSEN  rows 1-5 header, row 6 = group-header prototype, row 7 = person prototype,
                  signature block located by the cell "Diperiksa Oleh" (column G)
Last month's output also satisfies these anchors, so it can be used as the template too.
"""
import copy, io, re, collections, datetime as dt
import openpyxl
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.cell.cell import MergedCell

MONTHS_ID = ['', 'JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI', 'JULI', 'AGUSTUS',
             'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER']
STATUS_LABEL = {'S': 'Sakit', 'I': 'Ijin', 'C': 'Cuti', 'A': 'Tanpa berita'}
NCOL = 62           # MASTER DATA columns A..BJ
RK_NCOL = 44        # REKAP columns A..AR
MD_DAYS = 42        # MASTER DATA date columns F..AU
RK_DAYS = 31        # REKAP day columns D..AH


class TemplateError(ValueError):
    pass


def _v(ws, ref):
    v = ws[ref].value
    return v.strip() if isinstance(v, str) else v


def validate_template(wb):
    errs = []
    for name in ('MASTER DATA', '01 REKAP ABSEN'):
        if name not in wb.sheetnames:
            errs.append(f'sheet "{name}" is missing')
    if errs:
        raise TemplateError('Template layout not recognised: ' + '; '.join(errs) + '.')
    md, rk = wb['MASTER DATA'], wb['01 REKAP ABSEN']
    labels = [str(md.cell(7 + i, 5).value or '').strip().upper() for i in range(5)]
    if labels != ['IN', 'OUT', 'TT', 'LOKASI', 'STATUS']:
        errs.append(f'MASTER DATA E7:E11 should read IN / OUT / TT / Lokasi / Status (found {labels})')
    if not str(_v(md, 'B4') or '').lower().startswith('nama'):
        errs.append('MASTER DATA B4 should be the "Nama Personil" header')
    if not str(_v(rk, 'A6') or '').upper().startswith('ABSEN'):
        errs.append('01 REKAP ABSEN A6 should be a group header ("ABSEN <GRUP> :")')
    if not str(rk['B7'].value or '').startswith("='MASTER DATA'!B"):
        errs.append("01 REKAP ABSEN B7 should be a person row (='MASTER DATA'!B..)")
    if not any(str(rk.cell(r, 7).value or '').strip() == 'Diperiksa Oleh' for r in range(8, rk.max_row + 1)):
        errs.append('01 REKAP ABSEN has no "Diperiksa Oleh" signature block below row 7')
    if errs:
        raise TemplateError('Template layout not recognised: ' + '; '.join(errs) + '. '
                            'Use "FORMAT Timesheet - CLEAN.xlsx" or a file produced by this app.')


def _fill(c):
    return PatternFill(start_color='FF' + c, end_color='FF' + c, fill_type='solid')


def id_date(d, year=True):
    return f'{d.day} {MONTHS_ID[d.month].capitalize()}' + (f' {d.year}' if year else '')


def runs(dates):
    """Sorted dates -> list of (first, last) runs of consecutive days."""
    out = []
    for d in sorted(set(dates)):
        if out and (d - out[-1][1]).days == 1:
            out[-1] = (out[-1][0], d)
        else:
            out.append((d, d))
    return out


def fmt_run(a, b, year=False):
    """(a, b) -> '5 Oktober', '28-29 September', '30 September - 3 Oktober' (+ year if asked)."""
    y = f' {b.year}' if year else ''
    if a == b:
        return id_date(a, False) + y
    if (a.year, a.month) == (b.year, b.month):
        return f'{a.day}-{b.day} {MONTHS_ID[b.month].capitalize()}' + y
    return id_date(a, a.year != b.year and year) + ' - ' + id_date(b, False) + y


def ket_from_manual(entries):
    """[(date, status, note)] -> 'KET' text, e.g. '28 September, 10 Oktober Sakit (surat dokter)'."""
    groups = collections.OrderedDict()
    for d, st, note in sorted(entries):
        groups.setdefault((st, note or ''), []).append(d)
    lines = []
    for (st, note), ds in groups.items():
        parts = []
        for a, b in runs(ds):
            # merge single days of the same month: "1, 10, 14 Oktober"
            if a == b and parts and parts[-1][0] == 'days' and parts[-1][1] == (a.year, a.month):
                parts[-1][2].append(a.day)
            elif a == b:
                parts.append(['days', (a.year, a.month), [a.day]])
            else:
                parts.append(['run', None, fmt_run(a, b)])
        txt = ', '.join(', '.join(str(x) for x in p[2]) + ' ' + MONTHS_ID[p[1][1]].capitalize() if p[0] == 'days' else p[2]
                        for p in parts)
        lines.append(txt + ' ' + STATUS_LABEL.get(st.upper(), st) + (f' ({note})' if note else ''))
    return '\n'.join(lines)


def log_rows(manual, status, names):
    """Manual entries of one status -> rows (first_date, date text, name, note), consecutive days merged."""
    grp = collections.defaultdict(list)
    for m in manual:
        if m['status'].upper() == status:
            grp[(m['code'], (m.get('note') or '').strip())].append(dt.date.fromisoformat(m['date']))
    rows = []
    for (code, note), ds in grp.items():
        for a, b in runs(ds):
            rows.append((a, fmt_run(a, b, year=True), names.get(code, code), note))
    rows.sort(key=lambda r: (r[0], r[2]))
    return rows


def build(template_bytes, parsed, people, manual, opts=None):
    """people: ordered list of dict(code, name, grup, jab, sk, ket).
    manual: list of dict(code, date 'YYYY-MM-DD', status S/I/C/a, note).
    Returns (xlsx_bytes, report)."""
    opts = opts or {}
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    validate_template(wb)
    md, rk = wb['MASTER DATA'], wb['01 REKAP ABSEN']

    START = dt.date.fromisoformat(parsed['start']); END = dt.date.fromisoformat(parsed['end'])
    NDAYS = (END - START).days + 1
    if NDAYS > RK_DAYS:
        raise ValueError(f'The CSV covers {NDAYS} days ({START} to {END}); the timesheet has room for {RK_DAYS}. '
                         'Export one monthly period at a time.')
    dates = [START + dt.timedelta(i) for i in range(NDAYS)]

    def libur(d):
        return d.weekday() == 6          # Sunday

    man = collections.defaultdict(dict)
    for m in manual:
        man[m['code']][m['date']] = (m['status'].upper() if m['status'].upper() != 'A' else 'a', m.get('note') or '')

    report = collections.Counter()
    conflicts = []
    totals = collections.Counter()

    def status(code, d):
        rec = parsed['days'].get(f'{code}|{d}')
        mm = man.get(code, {}).get(str(d))
        st = None
        if rec and rec['out']:
            hrs = (dt.datetime.fromisoformat(rec['out']) - dt.datetime.fromisoformat(rec['in'])).total_seconds() / 3600
            if libur(d): st = 'L'
            elif rec['lok'] and 'WFH' in rec['lok'].upper(): st = 'W'
            elif hrs >= 6: st = 1
            else: st = 0.5
        elif rec:
            report['missing_out'] += 1
        if mm:
            if rec:
                conflicts.append(dict(code=code, date=str(d), recorded=st if st is not None else 'IN only',
                                      manual=mm[0], lok=rec.get('lok')))
            st = mm[0]
        return st, rec

    # ================= MASTER DATA =================
    BLOCK0 = 7
    blk_style = {(i, c): copy.copy(md.cell(BLOCK0 + i, c)._style) for i in range(5) for c in range(1, NCOL + 1)}
    blk_h = [md.row_dimensions[BLOCK0 + i].height for i in range(5)]
    keep_cf = [(str(cf.sqref), cf.rules) for cf in md.conditional_formatting
               if str(cf.sqref) in ('A1:E1', 'F6:BC6', 'F4:BJ4')]

    for m in [m for m in md.merged_cells.ranges if m.min_row >= 7]:
        md.unmerge_cells(str(m))
    md.tables.clear()
    if md.max_row >= 7:
        md.delete_rows(7, md.max_row)
    for r in list(md.row_dimensions):
        if r >= 7:
            md.row_dimensions[r].hidden = False

    for i in range(MD_DAYS):                         # date header F..AU
        d = START + dt.timedelta(i); c = 6 + i
        md.cell(4, c).value = dt.datetime.combine(d, dt.time()); md.cell(5, c).value = dt.datetime.combine(d, dt.time())
        md.cell(6, c).value = 'Libur' if libur(d) else 'Kerja'

    LASTC = L(5 + RK_DAYS)                           # totals cover F..AJ (= REKAP D..AH)
    row = 7; prev = None; md_rows = {}
    for q in people:
        r0 = row; md_rows[q['code']] = r0
        for i in range(5):
            md.row_dimensions[r0 + i].height = blk_h[i]
            for c in range(1, NCOL + 1):
                md.cell(r0 + i, c)._style = copy.copy(blk_style[(i, c)])
        md.cell(r0, 1).value = 1 if prev is None else f'=A{prev}+1'
        md.cell(r0, 2).value = q['name']; md.cell(r0, 3).value = q['jab']; md.cell(r0, 4).value = q['sk'] or None
        for i, lab in enumerate(['IN', 'OUT', 'TT', 'Lokasi', 'Status']):
            md.cell(r0 + i, 5).value = lab
        for j in range(MD_DAYS):
            c = 6 + j; col = L(c)
            md.cell(r0 + 2, c).value = f'={col}{r0+1}-{col}{r0}'
            if j < NDAYS:
                st, rec = status(q['code'], dates[j])
                if rec:
                    md.cell(r0, c).value = dt.datetime.fromisoformat(rec['in'])
                    if rec['out']:
                        md.cell(r0 + 1, c).value = dt.datetime.fromisoformat(rec['out'])
                    md.cell(r0 + 3, c).value = rec['lok']
                md.cell(r0 + 4, c).value = st
                if st is not None:
                    totals[str(st).upper()] += 1
        s = f'F{r0+4}:{LASTC}{r0+4}'
        for colname, f in zip(['AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC'],
                              [f'=SUMIF({s},1)+SUMIF({s},0.5)', f'=COUNTIF({s},"W")', f'=COUNTIF({s},"L")',
                               f'=COUNTIF({s},"s")', f'=COUNTIF({s},"i")', f'=COUNTIF({s},"C")', f'=COUNTIF({s},"a")']):
            md[f'{colname}{r0}'] = f
        md[f'BD{r0}'] = q.get('ket') or None
        for colname in ['A', 'B', 'C', 'D', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD']:
            md.merge_cells(f'{colname}{r0}:{colname}{r0+4}')
        prev = r0; row += 5
    md_end = row - 1

    md.conditional_formatting = ConditionalFormattingList()
    for rng, rules in keep_cf:
        for rule in rules:
            md.conditional_formatting.add(rng, rule)
    if people:
        for val, color in [('1', '70E096'), ('0.5', 'F4CCCC'), ('"W"', 'B6D7A8'), ('"L"', 'F9CB9C'), ('"S"', 'FFFF00'),
                           ('"I"', 'A4C2F4'), ('"C"', '99CEFF'), ('"a"', 'FF0000'), ('"TB"', 'FF0000')]:
            md.conditional_formatting.add(f'F7:AU{md_end}', CellIsRule(operator='equal', formula=[val], fill=_fill(color)))

    # ================= 01 REKAP ABSEN =================
    hdr_style = {c: copy.copy(rk.cell(6, c)._style) for c in range(1, RK_NCOL + 1)}
    per_style = {c: copy.copy(rk.cell(7, c)._style) for c in range(1, RK_NCOL + 1)}
    SIG0 = next(r for r in range(8, rk.max_row + 1) if str(rk.cell(r, 7).value or '').strip() == 'Diperiksa Oleh')
    sig = [[(rk.cell(SIG0 + i, c).value, copy.copy(rk.cell(SIG0 + i, c)._style)) for c in range(1, RK_NCOL + 1)]
           for i in range(6)]
    sig_h = [rk.row_dimensions[SIG0 + i].height for i in range(6)]
    h_hdr, h_per = rk.row_dimensions[6].height, rk.row_dimensions[7].height
    rk_cf = {str(cf.sqref): cf.rules for cf in rk.conditional_formatting}

    for m in [m for m in rk.merged_cells.ranges if m.min_row >= 6]:
        rk.unmerge_cells(str(m))
    rk.delete_rows(6, rk.max_row)
    for r in list(rk.row_dimensions):
        if r >= 6:
            rk.row_dimensions[r].hidden = False

    y0 = f' {START.year}' if START.year != END.year else ''
    rk['B1'] = f'ABSEN {START.day} {MONTHS_ID[START.month]}{y0}- {END.day} {MONTHS_ID[END.month]} {END.year}'
    rk['F5'] = "='MASTER DATA'!H6"
    # day header D..AH: formulas inside the period, blank after it (months of 28-31 days)
    for j in range(RK_DAYS):
        c = 4 + j
        if j < NDAYS:
            rk.cell(3, c).value = f'={L(c)}4'
            rk.cell(4, c).value = f"='MASTER DATA'!{L(6 + j)}5"
            rk.cell(5, c).value = f"='MASTER DATA'!{L(6 + j)}6"
        else:
            for r in (3, 4, 5):
                rk.cell(r, c).value = None

    row = 6; group = None; first = True
    for q in people:
        if q['grup'] != group:
            group = q['grup']; rk.row_dimensions[row].height = h_hdr
            for c in range(1, RK_NCOL + 1):
                rk.cell(row, c)._style = copy.copy(hdr_style[4 if 4 <= c <= 34 else c])
            rk.cell(row, 1).value = f'ABSEN {group.upper()} :'
            row += 1; first = True
        r0 = md_rows[q['code']]; rk.row_dimensions[row].height = h_per
        for c in range(1, RK_NCOL + 1):
            rk.cell(row, c)._style = copy.copy(per_style[c])
        rk.cell(row, 1).value = 1 if first else f'=A{row-1}+1'; first = False
        rk.cell(row, 2).value = f"='MASTER DATA'!B{r0}"; rk.cell(row, 3).value = f"='MASTER DATA'!C{r0}"
        for j in range(NDAYS):
            ref = f"'MASTER DATA'!{L(6 + j)}{r0 + 4}"
            rk.cell(row, 4 + j).value = f'=IF({ref}="","",{ref})'
        s = f'D{row}:AH{row}'
        for colname, f in zip(['AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP'],
                              [f'=SUMIF({s},1)+SUMIF({s},0.5)', f'=COUNTIF({s},"W")', f'=COUNTIF({s},"L")',
                               f'=COUNTIF({s},"s")', f'=COUNTIF({s},"i")', f'=COUNTIF({s},"C")', f'=COUNTIF({s},"a")']):
            rk[f'{colname}{row}'] = f
        rk[f'AQ{row}'] = f"=IF('MASTER DATA'!BD{r0}=\"\",\"\",'MASTER DATA'!BD{r0})"
        rk[f'AR{row}'] = f'=B{row}&CHAR(10)&AQ{row}'
        row += 1
    rk_end = row - 1
    row += 1
    for i in range(6):
        rk.row_dimensions[row + i].height = sig_h[i]
        for c in range(1, RK_NCOL + 1):
            v, stl = sig[i][c - 1]; cell = rk.cell(row + i, c); cell._style = stl
            if c < 43:
                cell.value = v
    sig_text = str(sig[0][3][0] or '')
    place = opts.get('place') or (sig_text.split(',')[0].strip() if ',' in sig_text else '')  # template's place, if any
    rk.cell(row, 4).value = f'{place}, {id_date(END)}' if place else id_date(END)

    rk.conditional_formatting = ConditionalFormattingList()
    for rule in rk_cf.get('D5:AH5', []):
        rk.conditional_formatting.add('D5:AH5', rule)
    blocks = []; start = None
    for rr in range(7, rk_end + 2):
        v = rk.cell(rr, 2).value
        is_person = rr <= rk_end and isinstance(v, str) and v.startswith('=')
        if is_person and start is None: start = rr
        if not is_person and start is not None: blocks.append((start, rr - 1)); start = None
    if blocks:
        sq = ' '.join(f'D{a}:AH{b}' for a, b in blocks)
        first_cell = f'D{blocks[0][0]}'
        for val, color in [('1', '6AA84F'), ('0.5', '93C47D'), ('"W"', 'B6D7A8'), ('"L"', 'F9CB9C'), ('"S"', 'FFFF00'),
                           ('"I"', 'A4C2F4'), ('"C"', '99CEFF'), ('"a"', 'FF0000'), ('"TB"', 'FF0000')]:
            rk.conditional_formatting.add(sq, CellIsRule(operator='equal', formula=[val], fill=_fill(color), stopIfTrue=True))
        rk.conditional_formatting.add(sq, FormulaRule(formula=[f'LEN(TRIM({first_cell}))=0'], fill=_fill('FFF2CC'), stopIfTrue=True))
        prio = 1
        for cf in rk.conditional_formatting:
            if str(cf.sqref) != 'D5:AH5':
                for rule in cf.rules: rule.priority = prio; prio += 1
        for cf in rk.conditional_formatting:
            if str(cf.sqref) == 'D5:AH5':
                for rule in cf.rules: rule.priority = prio; prio += 1

    # ================= other sheets =================
    periode = f'PERIODE {START.day} {MONTHS_ID[START.month]}{y0} - {END.day} {MONTHS_ID[END.month]} {END.year}'
    names = {q['code']: q['name'] for q in people}
    overflow = []
    if opts.get('clear_logs', True):
        _clear_logs(wb)
    if 'KET WFH, IZIN, DLL' in wb.sheetnames:
        kt = wb['KET WFH, IZIN, DLL']
        for r in range(1, kt.max_row + 1):
            v = kt.cell(r, 1).value
            if isinstance(v, str) and v.startswith('PERIODE'):
                kt.cell(r, 1).value = periode
    if opts.get('fill_logs', True):
        overflow = _fill_logs(wb, manual, names)

    wb.calculation.fullCalcOnLoad = True
    report['n'] = len(people)
    summary = dict(
        n=len(people), start=str(START), end=str(END), ndays=NDAYS,
        missing_out=report['missing_out'], conflicts=conflicts, log_overflow=overflow,
        totals=dict(hadir=totals['1'] + 0.5 * totals['0.5'], full=totals['1'], half=totals['0.5'],
                    wfh=totals['W'], lembur=totals['L'], sakit=totals['S'], izin=totals['I'],
                    cuti=totals['C'], tanpa_berita=totals['A']),
        md_end=md_end, rk_end=rk_end)
    return wb, summary


def _clear_logs(wb):
    """Remove the previous period's sick-note / WFH / izin / cuti logs (values only; keeps numbering)."""
    if '02.c SURAT SAKIT' in wb.sheetnames:
        ss = wb['02.c SURAT SAKIT']
        ss._images = []
        for r in range(6, ss.max_row + 1):
            for c in range(2, ss.max_column + 1):
                cell = ss.cell(r, c)
                if not isinstance(cell, MergedCell):
                    cell.value = None
    if 'KET WFH, IZIN, DLL' in wb.sheetnames:
        kt = wb['KET WFH, IZIN, DLL']
        kt._images = []
        for r in range(6, kt.max_row + 1):
            a = kt.cell(r, 1).value
            if isinstance(a, str) and a.strip() == 'No':
                continue
            for c in range(2, kt.max_column + 1):
                cell = kt.cell(r, c)
                if not isinstance(cell, MergedCell):
                    cell.value = None


def _section_rows(ws, header_text):
    """Data rows (numbered in column A) below the header row whose column D starts with header_text."""
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 4).value
        if isinstance(v, str) and v.strip().upper().startswith(header_text):
            rows = []
            rr = r + 1
            while rr <= ws.max_row and isinstance(ws.cell(rr, 1).value, (int, float)):
                rows.append(rr); rr += 1
            return rows
    return None


def _fill_logs(wb, manual, names):
    """Write this month's manual S / I / C entries into the log sheets. Returns overflow notes."""
    overflow = []
    sick = log_rows(manual, 'S', names)
    if sick and '02.c SURAT SAKIT' in wb.sheetnames:
        ss = wb['02.c SURAT SAKIT']
        proto = {c: copy.copy(ss.cell(15, c)._style) for c in range(1, 6)}
        for i, (_, when, name, note) in enumerate(sick):
            r = 6 + i
            if i >= 10:
                for c in range(1, 6):
                    ss.cell(r, c)._style = copy.copy(proto[c])
                ss.cell(r, 1).value = i + 1
            ss.cell(r, 2).value = when
            ss.cell(r, 3).value = name
            ss.cell(r, 4).value = 'Sakit'
            ss.cell(r, 5).value = note or None
    if 'KET WFH, IZIN, DLL' in wb.sheetnames:
        kt = wb['KET WFH, IZIN, DLL']
        for code, hdr, label in (('I', 'ALASAN IJIN', 'izin'), ('C', 'ALASAN CUTI', 'cuti')):
            ents = log_rows(manual, code, names)
            if not ents:
                continue
            rows = _section_rows(kt, hdr)
            if rows is None:
                overflow.append(f'No "{hdr}" table found in sheet KET WFH, IZIN, DLL; {len(ents)} {label} entries not listed there.')
                continue
            for r, (_, when, name, note) in zip(rows, ents):
                kt.cell(r, 2).value = when
                kt.cell(r, 3).value = name
                kt.cell(r, 4).value = note or STATUS_LABEL[code]
            if len(ents) > len(rows):
                extra = ents[len(rows):]
                overflow.append(f'The {label} table in "KET WFH, IZIN, DLL" has {len(rows)} rows; '
                                f'{len(extra)} more not listed there: ' + '; '.join(f'{n} {w}' for _, w, n, _ in extra))
    return overflow


def save(wb):
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()
