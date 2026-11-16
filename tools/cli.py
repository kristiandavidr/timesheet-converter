"""Run the converter locally (same Python code the web page runs).

  pip install openpyxl
  python tools/cli.py EXPORT.csv --prev "last month.xlsx" --manual manual.csv --out "Timesheet.xlsx"

manual.csv columns (same as the web page's export): Kode Anggota, Nama, Tanggal (YYYY-MM-DD or
DD/MM/YYYY), Sampai (optional end date), Status (S/I/C/a), Catatan. Nama is used only when
Kode Anggota is empty.
"""
import argparse, csv, json, os, sys, datetime as dt

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', 'py'))
import api  # noqa: E402


def _date(s):
    s = (s or '').strip()
    if not s:
        return None
    return dt.datetime.strptime(s, '%d/%m/%Y' if '/' in s else '%Y-%m-%d').date()


def read_manual(path, people):
    by_name = {p['name'].upper(): p['code'] for p in people}
    out = []
    with open(path, encoding='utf-8-sig') as f:
        for r in csv.DictReader(f):
            code = (r.get('Kode Anggota') or '').strip() or by_name.get((r.get('Nama') or '').strip().upper())
            d0, d1 = _date(r.get('Tanggal')), _date(r.get('Sampai'))
            if not (code and d0):
                continue
            d = d0
            while d <= (d1 or d0):
                out.append(dict(code=code, date=d.isoformat(), status=(r.get('Status') or '').strip(),
                                note=(r.get('Catatan') or '').strip()))
                d += dt.timedelta(days=1)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('csv')
    ap.add_argument('--prev', help="last month's output (roster source)")
    ap.add_argument('--template', default=os.path.join(HERE, '..', 'template', 'format-timesheet-clean.xlsx'))
    ap.add_argument('--manual')
    ap.add_argument('--out', default='Timesheet.xlsx')
    ap.add_argument('--use-waktu', action='store_true', help='use Tanggal+Waktu for entries added later')
    a = ap.parse_args()
    an = json.loads(api.load(a.csv, a.prev, a.use_waktu))
    print(f"period {an['period']['start']} .. {an['period']['end']}, {len(an['people'])} people, roster: {an['roster_source']}")
    for k in ('new_groups', 'not_in_export', 'duplicates'):
        if an[k]:
            print(k, json.dumps(an[k], ensure_ascii=False))
    print('new people:', [p['name'] for p in an['people'] if p['new']])
    print('moved:', [(p['name'], p['moved_from'], p['grup']) for p in an['people'] if p['moved_from']])
    manual = read_manual(a.manual, an['people']) if a.manual else []
    s = json.loads(api.convert(a.template, a.out, json.dumps(dict(manual=manual))))
    print(json.dumps({k: s[k] for k in ('n', 'start', 'end', 'missing_out', 'totals', 'log_overflow')}, ensure_ascii=False))
    for c in s['conflicts']:
        print('conflict', c)
    print('saved', a.out)


if __name__ == '__main__':
    main()
