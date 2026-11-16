"""The people roster, stored in a hidden `_ROSTER` sheet inside the output workbook.

Next month the user uploads last month's output; the roster (role, employment status, order,
group order, code aliases) comes with it, so no personal data has to live in the repo.

If the uploaded workbook has no `_ROSTER` sheet (e.g. a file made before this app existed), the
roster is seeded from its MASTER DATA sheet: entries then have a name but no Kode Anggota, and
are matched to this month's people by name.
"""
import re, difflib

SHEET = '_ROSTER'
COLS = ['Kode Anggota', 'Nama', 'Grup', 'Jabatan', 'Status Karyawan', 'Active (Y/N)',
        'First seen', 'Last seen', 'Urutan', 'Alias (kode lama)']
GROUP_COL = 12          # column L: group order


def norm_name(s):
    s = (s or '').upper()
    s = re.sub(r'[^A-Z0-9 ]+', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def name_score(a, b):
    """0..1 similarity between two person names."""
    a, b = norm_name(a), norm_name(b)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    ta, tb = a.split(), b.split()
    short, long_ = (ta, tb) if len(ta) <= len(tb) else (tb, ta)
    if len(short) >= 2 and all(t in long_ for t in short):
        return 0.95
    r = difflib.SequenceMatcher(None, a, b).ratio()
    r2 = difflib.SequenceMatcher(None, ' '.join(ta[:2]), ' '.join(tb[:2])).ratio() if len(ta) > 1 and len(tb) > 1 else 0
    return max(r, r2 * 0.95)


def _s(v):
    if v is None:
        return ''
    if hasattr(v, 'strftime'):
        return v.strftime('%Y-%m-%d')
    return str(v).strip()


def split_jabatan(cell):
    """MASTER DATA 'Jabatan/Proyek' cell -> (role, grup)."""
    s = _s(cell)
    if not s:
        return '', ''
    m = re.search(r'Proyek\s*:?\s*(.*)$', s, flags=re.S)
    grup = re.sub(r'\s+', ' ', m.group(1)).strip() if m else ''
    role = s[:m.start()] if m else s
    role = re.sub(r'\s+', ' ', role).strip().rstrip(':').strip()
    return role, grup


def read_roster(wb):
    """-> (entries, group_order, source). source: 'roster' | 'master' | 'none'."""
    if SHEET in wb.sheetnames:
        ws = wb[SHEET]
        hdr = [_s(ws.cell(1, c).value) for c in range(1, len(COLS) + 1)]
        if hdr[:3] == COLS[:3]:
            entries = []
            for r in range(2, ws.max_row + 1):
                v = [ws.cell(r, c).value for c in range(1, len(COLS) + 1)]
                if not any(x not in (None, '') for x in v):
                    continue
                try:
                    order = float(v[8]) if v[8] not in (None, '') else None
                except (TypeError, ValueError):
                    order = None
                entries.append(dict(
                    code=_s(v[0]), nama=_s(v[1]), grup=_s(v[2]), jabatan=_s(v[3]), sk=_s(v[4]),
                    active=_s(v[5]).upper() != 'N', first_seen=_s(v[6]), last_seen=_s(v[7]), order=order,
                    alias=[a.strip() for a in _s(v[9]).split(',') if a.strip()]))
            groups = []
            for r in range(2, ws.max_row + 1):
                g = _s(ws.cell(r, GROUP_COL).value)
                if g and g not in groups:
                    groups.append(g)
            return entries, groups, 'roster'

    # ---- seed from MASTER DATA (no codes, matched by name) ----
    if 'MASTER DATA' not in wb.sheetnames:
        return [], [], 'none'
    md = wb['MASTER DATA']
    entries = []
    for r in range(7, md.max_row + 1):
        if _s(md.cell(r, 5).value).upper() == 'IN' and _s(md.cell(r, 2).value) and not _s(md.cell(r, 2).value).startswith('='):
            role, grup = split_jabatan(md.cell(r, 3).value)
            sk = _s(md.cell(r, 4).value).replace('\U0001F920', '').strip()
            entries.append(dict(code='', nama=re.sub(r'\s+', ' ', _s(md.cell(r, 2).value)).upper(), grup=grup,
                                jabatan=role, sk=sk, active=True, first_seen='', last_seen='',
                                order=float(len(entries) + 1), alias=[]))
    groups = []
    if '01 REKAP ABSEN' in wb.sheetnames:
        rk = wb['01 REKAP ABSEN']
        for r in range(6, rk.max_row + 1):
            m = re.match(r'\s*ABSEN\s+(.*?)\s*:\s*$', _s(rk.cell(r, 1).value))
            if m and '<' not in m.group(1) and m.group(1) not in groups:
                groups.append(m.group(1))
    if not groups:
        for e in entries:
            if e['grup'] and e['grup'] not in groups:
                groups.append(e['grup'])
    return entries, groups, ('master' if entries else 'none')


def alias_map(entries):
    m = {}
    for e in entries:
        for a in e.get('alias', []):
            if e['code'] and a and a != e['code']:
                m[a] = e['code']
    return m


DUP_THRESHOLD = 0.82


def reconcile(entries, groups, parsed):
    """Match this month's people (from the CSV) to the roster.

    Returns dict(people=[...], groups=[...], new_groups=[...], not_in_export=[...], duplicates=[...]).
    Each person: code, name, grup, jabatan, sk, order, rid (roster index or None), flags.
    """
    by_code = {e['code']: i for i, e in enumerate(entries) if e['code']}
    used = set()
    people = []
    csv_codes = {p['code'] for p in parsed['people']}
    for p in parsed['people']:
        rid = by_code.get(p['code'])
        how = 'code' if rid is not None else None
        if rid is None:                           # seeded entries have no code: exact name match
            nn = norm_name(p['name'])
            for i, e in enumerate(entries):
                if i not in used and not e['code'] and norm_name(e['nama']) == nn:
                    rid, how = i, 'name'
                    break
        if rid is not None:
            used.add(rid)
        e = entries[rid] if rid is not None else None
        person = dict(code=p['code'], name=p['name'], grup=p['grup'], grups=p['grups'],
                      codes=p.get('codes', [p['code']]),
                      jabatan=e['jabatan'] if e else '', sk=e['sk'] if e else '',
                      order=e['order'] if e else None, rid=rid, matched_by=how,
                      new=e is None, moved_from=None,
                      multi_group=len(p['grups']) > 1)
        if e and e['grup'] and e['grup'].upper() != p['grup'].upper():
            person['moved_from'] = e['grup']
        people.append(person)

    # roster entries not in this export
    not_in = []
    for i, e in enumerate(entries):
        if i in used:
            continue
        if e['code'] and e['code'] in csv_codes:
            continue
        if e.get('active', True) or not e['code']:
            not_in.append(dict(rid=i, code=e['code'], name=e['nama'], grup=e['grup']))

    # possible duplicates: a new code whose name looks like someone already known
    dups = []
    for p in people:
        if not p['new']:
            continue
        cands = []
        for i, e in enumerate(entries):
            sc = name_score(p['name'], e['nama'])
            if sc >= DUP_THRESHOLD:
                cands.append(dict(kind='roster', rid=i, code=e['code'], name=e['nama'], grup=e['grup'],
                                  score=round(sc, 2), in_export=bool(e['code'] and e['code'] in csv_codes)))
        for q in people:
            if q is p:
                continue
            sc = name_score(p['name'], q['name'])
            if sc >= DUP_THRESHOLD and q['rid'] is None:
                cands.append(dict(kind='csv', rid=None, code=q['code'], name=q['name'], grup=q['grup'],
                                  score=round(sc, 2), in_export=True))
        if cands:
            cands.sort(key=lambda c: -c['score'])
            dups.append(dict(code=p['code'], name=p['name'], grup=p['grup'], candidates=cands))
    # a CSV/CSV pair shows up twice (A~B and B~A); keep one
    seen, uniq = set(), []
    for d in dups:
        cands = [c for c in d['candidates'] if not (c['kind'] == 'csv' and (c['code'], d['code']) in seen)]
        for c in cands:
            if c['kind'] == 'csv':
                seen.add((d['code'], c['code']))
        if cands:
            uniq.append(dict(d, candidates=cands))

    # group order: roster/template order (case-insensitive), then new groups in CSV order
    csv_groups = []
    for p in parsed['people']:
        if p['grup'] not in csv_groups:
            csv_groups.append(p['grup'])
    upper = {g.upper(): g for g in csv_groups}
    order, known_upper = [], set()
    for g in groups:
        known_upper.add(g.upper())
        real = upper.get(g.upper(), g)
        if real not in order:
            order.append(real)
    new_groups = sorted((g for g in csv_groups if g.upper() not in known_upper), key=str.upper)
    order += new_groups
    return dict(people=people, groups=order, new_groups=new_groups if groups else [],
                not_in_export=not_in, duplicates=uniq)


def sort_people(people, group_order):
    gi = {g.upper(): i for i, g in enumerate(group_order)}
    return sorted(people, key=lambda q: (gi.get(q['grup'].upper(), len(gi)),
                                         q['order'] if q.get('order') is not None else 1e9,
                                         norm_name(q['name'])))


def write_roster(wb, entries, group_order):
    if SHEET in wb.sheetnames:
        del wb[SHEET]
    ws = wb.create_sheet(SHEET)
    for c, h in enumerate(COLS, start=1):
        ws.cell(1, c).value = h
    ws.cell(1, GROUP_COL).value = 'Urutan Grup'
    ws.cell(1, GROUP_COL + 2).value = ('Used by the Timesheet Converter. Upload this file next month as '
                                       '"last month\'s output" so roles, status and order carry over.')
    for r, e in enumerate(entries, start=2):
        vals = [e['code'], e['nama'], e['grup'], e['jabatan'], e['sk'], 'Y' if e['active'] else 'N',
                e['first_seen'], e['last_seen'], e['order'], ', '.join(e.get('alias', []))]
        for c, v in enumerate(vals, start=1):
            ws.cell(r, c).value = v if v not in ('', None) else None
    for r, g in enumerate(group_order, start=2):
        ws.cell(r, GROUP_COL).value = g
    widths = [14, 32, 28, 36, 28, 10, 12, 12, 8, 18, 2, 30]
    from openpyxl.utils import get_column_letter
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.sheet_state = 'hidden'
    return ws
