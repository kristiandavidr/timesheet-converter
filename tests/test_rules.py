"""Rule tests on a small made-up export (no real people). Run: python tests/test_rules.py"""
import csv, io, json, os, sys, tempfile, datetime as dt

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, '..')
sys.path.insert(0, os.path.join(ROOT, 'py'))
import openpyxl  # noqa: E402
import api, parse, build  # noqa: E402

COLS = ['Tanggal', 'Nama Lengkap', 'Kode Anggota', 'Grup', 'EntryType', 'Waktu', 'Aktivitas', 'Proyek', 'Dibuat pada', 'Pencatat']


def row(code, name, grup, typ, when, act='', created=None):
    t = dt.datetime.strptime(when, '%Y-%m-%d %H:%M:%S')
    c = dt.datetime.strptime(created, '%Y-%m-%d %H:%M:%S') if created else t
    return {'Tanggal': t.strftime('%d/%m/%Y'), 'Nama Lengkap': name, 'Kode Anggota': code, 'Grup': grup,
            'EntryType': typ, 'Waktu': t.strftime('%I:%M %p'), 'Aktivitas': act, 'Proyek': '',
            'Dibuat pada': c.strftime('%d/%m/%Y %H.%M.%S +07:00'), 'Pencatat': name}


A = ('T-1', 'Tester Satu', 'Kantor Uji')
B = ('T-2', 'Tester Dua', 'Proyek Uji')
rows = [
    # 2026-03-02 (Mon): full day at the office -> 1
    row(*A, 'In', '2026-03-02 08:00:00', 'Kantor'), row(*A, 'Out', '2026-03-02 17:00:00'),
    # 03-03: 4 hours -> 0.5
    row(*A, 'In', '2026-03-03 08:00:00', 'Kantor'), row(*A, 'Out', '2026-03-03 12:00:00'),
    # 03-04: WFA -> WFH -> W
    row(*A, 'In', '2026-03-04 08:00:00', 'WFA'), row(*A, 'Out', '2026-03-04 16:00:00'),
    # 03-05: forgot to clock out; Out + In next morning within 5 min -> blank on 03-05, normal 03-06
    row(*A, 'In', '2026-03-05 08:00:00', 'Kantor'), row(*A, 'Out', '2026-03-06 07:58:00'),
    row(*A, 'In', '2026-03-06 07:59:00', 'Kantor'), row(*A, 'Out', '2026-03-06 16:00:00'),
    # 03-08 (Sunday) with attendance -> L
    row(*A, 'In', '2026-03-08 09:00:00', 'Lokasi Proyek'), row(*A, 'Out', '2026-03-08 15:00:00'),
    # 03-09: break splits the day; IN 08:00, OUT 17:00 -> 1; Lokasi falls back to Grup
    row(*A, 'In', '2026-03-09 08:00:00'), row(*A, 'StartBreak', '2026-03-09 12:00:00'),
    row(*A, 'In', '2026-03-09 13:00:00', 'Lokasi Proyek'), row(*A, 'Out', '2026-03-09 17:00:00'),
    # B: overnight shift counts on the start day
    row(*B, 'In', '2026-03-02 20:00:00', 'Lokasi Proyek'), row(*B, 'Out', '2026-03-03 06:00:00'),
    # B: session longer than 24h -> blank
    row(*B, 'In', '2026-03-10 08:00:00', 'Lokasi Proyek'), row(*B, 'Out', '2026-03-11 10:00:00'),
    # B: punch added two days later by a manager (backdated)
    row(*B, 'In', '2026-03-12 08:00:00', 'Lokasi Proyek', created='2026-03-14 10:00:00'),
    row(*B, 'Out', '2026-03-12 17:00:00', created='2026-03-14 10:01:00'),
    # B: last day of the period, clock-in only -> blank
    row(*B, 'In', '2026-03-31 08:00:00', 'Lokasi Proyek'),
]


def write_csv(path, rows):
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=COLS); w.writeheader(); w.writerows(rows)


def status_cells(path):
    wb = openpyxl.load_workbook(path)
    md = wb['MASTER DATA']
    out = {}
    for r in range(7, md.max_row + 1, 5):
        name = md.cell(r, 2).value
        for c in range(6, 6 + 31):
            d = md.cell(4, c).value.date().isoformat()
            out[(name, d)] = (md.cell(r + 4, c).value, md.cell(r + 3, c).value)
    return wb, out


def main():
    tmp = tempfile.mkdtemp()
    f = os.path.join(tmp, 'export.csv'); write_csv(f, rows)
    tpl = os.path.join(ROOT, 'template', 'format-timesheet-clean.xlsx')
    out = os.path.join(tmp, 'out.xlsx')

    an = json.loads(api.load(f))
    assert an['period'] == {'start': '2026-03-02', 'end': '2026-03-31', 'ndays': 30}, an['period']
    assert [p['code'] for p in an['people']] == ['T-1', 'T-2'] or len(an['people']) == 2
    manual = [{'code': 'T-1', 'date': '2026-03-03', 'status': 'S', 'note': 'surat dokter'},
              {'code': 'T-1', 'date': '2026-03-11', 'status': 'I', 'note': ''},
              {'code': 'T-1', 'date': '2026-03-12', 'status': 'I', 'note': ''}]
    s = json.loads(api.convert(tpl, out, json.dumps({'manual': manual, 'people': {'T-2': {'jabatan': 'Pengawas'}}})))
    wb, st = status_cells(out)
    T1, T2 = 'TESTER SATU', 'TESTER DUA'
    expect = {
        (T1, '2026-03-02'): (1, 'KANTOR'),
        (T1, '2026-03-03'): ('S', 'KANTOR'),           # manual S overrides 0.5 (and is reported as a clash)
        (T1, '2026-03-04'): ('W', 'WFH'),
        (T1, '2026-03-05'): (None, 'KANTOR'),          # forgotten clock-out
        (T1, '2026-03-06'): (1, 'KANTOR'),
        (T1, '2026-03-08'): ('L', 'LOKASI PROYEK'),
        (T1, '2026-03-09'): (1, 'KANTOR UJI> LOKASI PROYEK'),
        (T1, '2026-03-11'): ('I', None),
        (T2, '2026-03-02'): (1, 'LOKASI PROYEK'),      # overnight, counted on start day
        (T2, '2026-03-03'): (None, None),
        (T2, '2026-03-10'): (None, 'LOKASI PROYEK'),   # > 24h
        (T2, '2026-03-14'): (0.5, 'LOKASI PROYEK'),    # backdated punches count when created (default): 1 minute
        (T2, '2026-03-31'): (None, 'LOKASI PROYEK'),   # IN only
    }
    for k, v in expect.items():
        assert st[k] == v, (k, st[k], v)
    assert s['totals']['sakit'] == 1 and s['totals']['izin'] == 2 and s['totals']['lembur'] == 1
    assert s['missing_out'] == 3, s['missing_out']
    assert [c['date'] for c in s['conflicts']] == ['2026-03-03'], s['conflicts']
    md = wb['MASTER DATA']
    assert md['C12'].value == 'Pengawas\nProyek : Proyek Uji', md['C12'].value
    assert md['BD7'].value == '3 Maret Sakit (surat dokter)\n11-12 Maret Ijin', md['BD7'].value
    rk = wb['01 REKAP ABSEN']
    assert rk['B1'].value == 'ABSEN 2 MARET- 31 MARET 2026'
    assert rk['AG3'].value == '=AG4' and rk['AH3'].value is None      # 30-day period: AH cleared
    assert wb.calculation.fullCalcOnLoad
    assert wb['_ROSTER'].sheet_state == 'hidden'
    assert any(str(rk.cell(r, 4).value).endswith('31 Maret 2026') for r in range(1, rk.max_row + 1))
    assert wb['KET WFH, IZIN, DLL']['B17'].value == '11-12 Maret 2026'

    # use_waktu: the backdated punch moves to the day it was meant for
    an = json.loads(api.set_option(True))
    s = json.loads(api.convert(tpl, out, json.dumps({'people': {'T-2': {'jabatan': 'Pengawas'}}})))
    _, st = status_cells(out)
    assert st[(T2, '2026-03-12')] == (1, 'LOKASI PROYEK') and st[(T2, '2026-03-14')] == (None, None)

    # next month: last output as the roster source keeps the Jabatan
    api.load(f, out)
    an = json.loads(api._analyze())
    assert an['roster_source'] == 'roster'
    assert next(p for p in an['people'] if p['code'] == 'T-2')['jabatan'] == 'Pengawas'
    assert not any(p['new'] for p in an['people'])

    # 31-day period fills AH
    write_csv(f, rows + [row(*A, 'In', '2026-04-01 08:00:00', 'Kantor'), row(*A, 'Out', '2026-04-01 17:00:00')])
    api.load(f)
    json.loads(api.convert(tpl, out, json.dumps({})))
    rk = openpyxl.load_workbook(out)['01 REKAP ABSEN']
    assert rk['AH3'].value == '=AH4'

    try:
        parse.read_entries(b'Hari,Tanggal\n1,2\n')
        raise AssertionError('expected CsvError')
    except parse.CsvError:
        pass
    print('all rule tests passed')


if __name__ == '__main__':
    main()
