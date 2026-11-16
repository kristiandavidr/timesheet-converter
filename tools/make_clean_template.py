"""Make a public-safe template from the real FORMAT Timesheet.xlsx:
no people, no notes, no sick-letter images, no old-period data.
Layout anchors (used by build.py):
  MASTER DATA      rows 1-6 header, rows 7-11 = one blank prototype person block
  01 REKAP ABSEN   rows 1-5 header, row 6 = group header prototype, row 7 = person prototype,
                   rows 9-14 = signature block
"""
import sys, copy
import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.utils import get_column_letter as L

SRC, OUT = sys.argv[1:3]
wb = openpyxl.load_workbook(SRC)
md = wb['MASTER DATA']; rk = wb['01 REKAP ABSEN']

# ---------- MASTER DATA: keep header, one blank block styled like rows 22-26 ----------
blk = {(i, c): copy.copy(md.cell(22 + i, c)._style) for i in range(5) for c in range(1, 63)}
blk_h = [md.row_dimensions[22 + i].height for i in range(5)]
hdr_cf = [(str(cf.sqref), cf.rules) for cf in md.conditional_formatting if str(cf.sqref) in ('A1:E1', 'F6:BC6', 'F4:BJ4')]
for m in [m for m in md.merged_cells.ranges if m.min_row >= 7]: md.unmerge_cells(str(m))
md.tables.clear()
md.delete_rows(7, md.max_row)
for r in list(md.row_dimensions):
    if r >= 7: md.row_dimensions[r].hidden = False
for i in range(5):
    md.row_dimensions[7 + i].height = blk_h[i]
    for c in range(1, 63): md.cell(7 + i, c)._style = copy.copy(blk[(i, c)])
for i, lab in enumerate(['IN', 'OUT', 'TT', 'Lokasi', 'Status']): md.cell(7 + i, 5).value = lab
md['A7'] = 1
for c in range(6, 48): md.cell(9, c).value = f'={L(c)}8-{L(c)}7'
s = 'F11:AJ11'
for col, f in zip(['AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC'],
    [f'=SUMIF({s},1)+SUMIF({s},0.5)', f'=COUNTIF({s},"W")', f'=COUNTIF({s},"L")', f'=COUNTIF({s},"s")',
     f'=COUNTIF({s},"i")', f'=COUNTIF({s},"C")', f'=COUNTIF({s},"a")']):
    md[f'{col}7'] = f
for col in ['A', 'B', 'C', 'D', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD']: md.merge_cells(f'{col}7:{col}11')
md.conditional_formatting = ConditionalFormattingList()
for rng, rules in hdr_cf:
    for rule in rules: md.conditional_formatting.add(rng, rule)

# ---------- 01 REKAP ABSEN ----------
hdr = {c: copy.copy(rk.cell(6, c)._style) for c in range(1, 45)}
per = {c: copy.copy(rk.cell(7, c)._style) for c in range(1, 45)}
sig = [[(rk.cell(72 + i, c).value, copy.copy(rk.cell(72 + i, c)._style)) for c in range(1, 45)] for i in range(6)]
sig_h = [rk.row_dimensions[72 + i].height for i in range(6)]
h6, h7 = rk.row_dimensions[6].height, rk.row_dimensions[7].height
d5_cf = [(str(cf.sqref), cf.rules) for cf in rk.conditional_formatting if str(cf.sqref).startswith('D5:V5')]
for m in [m for m in rk.merged_cells.ranges if m.min_row >= 6]: rk.unmerge_cells(str(m))
rk.delete_rows(6, rk.max_row)
for r in list(rk.row_dimensions):
    if r >= 6: rk.row_dimensions[r].hidden = False
rk['B1'] = 'ABSEN (periode diisi otomatis)'
rk['F5'] = "='MASTER DATA'!H6"
rk.row_dimensions[6].height = h6; rk.row_dimensions[7].height = h7
for c in range(1, 45):
    rk.cell(6, c)._style = copy.copy(hdr[4 if 4 <= c <= 34 else c])
    rk.cell(7, c)._style = copy.copy(per[c])
rk['A6'] = 'ABSEN <GRUP> :'
rk['A7'] = 1; rk['B7'] = "='MASTER DATA'!B7"; rk['C7'] = "='MASTER DATA'!C7"
for j in range(31):
    ref = f"'MASTER DATA'!{L(6 + j)}11"
    rk.cell(7, 4 + j).value = f'=IF({ref}="","",{ref})'
s = 'D7:AH7'
for col, f in zip(['AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP'],
    [f'=SUMIF({s},1)+SUMIF({s},0.5)', f'=COUNTIF({s},"W")', f'=COUNTIF({s},"L")', f'=COUNTIF({s},"s")',
     f'=COUNTIF({s},"i")', f'=COUNTIF({s},"C")', f'=COUNTIF({s},"a")']):
    rk[f'{col}7'] = f
rk['AQ7'] = "=IF('MASTER DATA'!BD7=\"\",\"\",'MASTER DATA'!BD7)"
rk['AR7'] = '=B7&CHAR(10)&AQ7'
for i in range(6):
    rk.row_dimensions[9 + i].height = sig_h[i]
    for c in range(1, 45):
        v, stl = sig[i][c - 1]; cell = rk.cell(9 + i, c); cell._style = stl
        cell.value = v if c < 43 else None
rk.conditional_formatting = ConditionalFormattingList()
for rng, rules in d5_cf:
    for rule in rules: rk.conditional_formatting.add('D5:AH5', rule)

# ---------- other sheets ----------
for name in [n for n in wb.sheetnames if n not in ('01 REKAP ABSEN', 'MASTER DATA', '02.c SURAT SAKIT', 'KET WFH, IZIN, DLL')]:
    del wb[name]
ss = wb['02.c SURAT SAKIT']
ss._images = []
for r in range(6, ss.max_row + 1):
    for c in range(1, ss.max_column + 1):
        cell = ss.cell(r, c)
        if not isinstance(cell, openpyxl.cell.cell.MergedCell): cell.value = None
for i in range(10): ss.cell(6 + i, 1).value = i + 1
kt = wb['KET WFH, IZIN, DLL']
for r in range(1, kt.max_row + 1):
    v = kt.cell(r, 1).value
    if isinstance(v, str) and v.startswith('PERIODE'): kt.cell(r, 1).value = 'PERIODE (diisi otomatis)'
    for c in range(2, kt.max_column + 1):
        cell = kt.cell(r, c)
        hdr_row = isinstance(kt.cell(r, 1).value, str) and kt.cell(r, 1).value == 'No'
        if not hdr_row and not isinstance(cell, openpyxl.cell.cell.MergedCell) and r > 5: cell.value = None

wb.properties.creator = None; wb.properties.lastModifiedBy = None
wb.save(OUT)
print('saved', OUT)


# ---------- drop leftover custom table styles (unused, and their names come from old sheets) ----------
import re, zipfile, shutil, tempfile, os
tmp = tempfile.mktemp(suffix='.xlsx')
with zipfile.ZipFile(OUT) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'xl/styles.xml':
            s = data.decode('utf-8')
            s = re.sub(r'<tableStyle\b[^>]*?(/>|>.*?</tableStyle>)', '', s, flags=re.S)
            s = re.sub(r'(<tableStyles\b[^>]*\bcount=")\d+(")', r'\g<1>0\2', s)
            data = s.encode('utf-8')
        zout.writestr(item, data)
shutil.move(tmp, OUT)
print('stripped custom table styles')
