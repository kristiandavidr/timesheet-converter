"""Jibble "Monthly Raw Time Entries" CSV -> one IN / OUT / Lokasi per person per day.

Rules (see HANDOFF.md):
  * Punch time = "Dibuat pada" (has seconds), local wall-clock time of the punch.
    Optionally (use_waktu=True) entries that were added later by someone (Dibuat pada more than
    BACKDATE_MIN minutes away from Tanggal + Waktu) use Tanggal + Waktu instead.
  * Punches are ordered per person by that time. An "In" while already clocked in is an activity
    switch (same session), unless the open session started more than 24h earlier.
    "StartBreak" ends a session like "Out"; the next "In" starts a new one.
  * The day's IN is the first session start, OUT is the last valid clock-out of that day's sessions.
  * A session counts on the day its first "In" happened (overnight shifts count on the start day).
  * An "Out" followed within 5 minutes by an "In" on a later day than the session start is a
    forgotten clock-out: that session's OUT is missing. Sessions longer than 24h are missing too.
  * Lokasi = Aktivitas of each "In" of the day (WFA -> WFH), or Grup when Aktivitas is empty,
    upper-cased, consecutive duplicates collapsed, joined with "> ".
"""
import csv, io, re, collections, datetime as dt

FORGOT_GAP_MIN = 5
MAX_SESSION_H = 24
BACKDATE_MIN = 60

REQUIRED_COLS = ['Tanggal', 'Nama Lengkap', 'Kode Anggota', 'Grup', 'EntryType', 'Waktu', 'Dibuat pada']


class CsvError(ValueError):
    pass


def clean(s):
    return re.sub(r'\s+', ' ', s).strip() if s else ''


def _parse_created(s):
    # "26/09/2026 07.19.18 +07:00" -> (naive local datetime, aware datetime)
    m = re.match(r'\s*(\d{1,2})/(\d{1,2})/(\d{4})\s+(\d{1,2})[.:](\d{2})[.:](\d{2})\s*([+-])(\d{2}):?(\d{2})', s or '')
    if not m:
        return None, None
    d, mo, y, h, mi, se, sign, oh, om = m.groups()
    local = dt.datetime(int(y), int(mo), int(d), int(h), int(mi), int(se))
    off = dt.timedelta(hours=int(oh), minutes=int(om)) * (1 if sign == '+' else -1)
    return local, local.replace(tzinfo=dt.timezone(off))


def _parse_waktu(tanggal, waktu):
    for fmt in ('%d/%m/%Y %I:%M %p', '%d/%m/%Y %H:%M'):
        try:
            return dt.datetime.strptime(f'{tanggal.strip()} {waktu.strip()}', fmt)
        except ValueError:
            pass
    return None


def _activity(row):
    a = clean(row.get('Aktivitas'))
    if a.upper() == 'WFA':
        a = 'WFH'
    if not a:
        a = clean(row.get('Grup'))
    return a.upper() or None


def read_entries(data):
    """bytes/str -> list of punch dicts (all EntryTypes), plus the list of CSV dates."""
    if isinstance(data, (bytes, bytearray, memoryview)):
        data = bytes(data).decode('utf-8-sig')
    rd = csv.DictReader(io.StringIO(data))
    cols = rd.fieldnames or []
    missing = [c for c in REQUIRED_COLS if c not in cols]
    if missing:
        raise CsvError('This does not look like a Jibble "Monthly Raw Time Entries" export. '
                       'Missing column(s): ' + ', '.join(missing))
    entries, tanggal = [], set()
    for n, row in enumerate(rd, start=2):
        try:
            tg = dt.datetime.strptime(row['Tanggal'].strip(), '%d/%m/%Y').date()
        except ValueError:
            continue
        tanggal.add(tg)
        local, aware = _parse_created(row['Dibuat pada'])
        waktu = _parse_waktu(row['Tanggal'], row['Waktu'])
        if local is None and waktu is None:
            continue
        if local is None:                       # no "Dibuat pada": fall back to Tanggal + Waktu
            local, aware = waktu, waktu.replace(tzinfo=dt.timezone(dt.timedelta(hours=7)))
        entries.append(dict(
            line=n, code=clean(row['Kode Anggota']), name=clean(row['Nama Lengkap']), grup=clean(row['Grup']),
            type=clean(row['EntryType']), local=local, aware=aware, waktu=waktu,
            act=_activity(row), tanggal=tg,
            backdated=(waktu is not None and abs((local - waktu).total_seconds()) > BACKDATE_MIN * 60),
            by=clean(row.get('Pencatat'))))
    if not entries:
        raise CsvError('The CSV has no time entries.')
    return entries, sorted(tanggal)


def build_days(entries, alias=None, use_waktu=False):
    """Group punches into sessions and days.

    alias: {old_code: new_code}, applied before grouping (merged / re-created Jibble accounts).
    Returns dict(people=[...], days={'CODE|YYYY-MM-DD': {...}}, issues={...}).
    """
    alias = alias or {}
    per = collections.defaultdict(list)
    for e in entries:
        code = alias.get(e['code'], e['code'])
        t = e['local']
        if use_waktu and e['backdated'] and e['waktu'] is not None:
            t = e['waktu']
        per[code].append(dict(e, orig=e['code'], code=code, t=t, sort=(t if use_waktu and e['backdated'] else e['aware'])))

    people, days = [], {}
    issues = dict(missing_out=[], long_shifts=[], backdated=[], multi_group=[])
    for code, es in per.items():
        # sort by absolute time; entries placed by Tanggal+Waktu sort by their naive time in UTC+7
        def key(x):
            s = x['sort']
            if s.tzinfo is None:
                s = s.replace(tzinfo=dt.timezone(dt.timedelta(hours=7)))
            return (s, x['line'])
        es.sort(key=key)
        names = collections.Counter(x['name'] for x in es)
        last_name = es[-1]['name']
        grups = collections.Counter(x['grup'] for x in es if x['grup'])
        grup = grups.most_common(1)[0][0] if grups else ''
        if len(grups) > 1:
            issues['multi_group'].append(dict(code=code, name=last_name, groups=dict(grups), used=grup))
        people.append(dict(code=code, name=last_name, names=sorted(names), grup=grup, grups=dict(grups),
                           codes=sorted({x['orig'] for x in es})))

        # ---- sessions ----
        sessions, cur = [], None
        punches = [x for x in es if x['type'] in ('In', 'Out', 'StartBreak')]
        for i, x in enumerate(punches):
            if x['type'] == 'In':
                if cur is not None and x['t'] - cur['start'] > dt.timedelta(hours=MAX_SESSION_H):
                    cur['missing'] = True; cur['why'] = 'no clock-out'   # still "in" from days ago
                    cur = None
                if cur is None:
                    cur = dict(start=x['t'], end=None, acts=[], missing=False, why=None)
                    sessions.append(cur)
                if x['act']:
                    cur['acts'].append(x['act'])
            else:  # Out, or StartBreak (a break also ends the session; the next In starts a new one)
                if cur is None:
                    continue                    # stray Out without an In
                cur['end'] = x['t']
                nxt = punches[i + 1] if i + 1 < len(punches) else None
                if (nxt is not None and nxt['type'] == 'In'
                        and dt.timedelta(0) <= (nxt['t'] - x['t']) <= dt.timedelta(minutes=FORGOT_GAP_MIN)
                        and nxt['t'].date() > cur['start'].date()):
                    cur['missing'] = True; cur['why'] = 'forgotten clock-out (Out + In next day)'
                elif cur['end'] - cur['start'] > dt.timedelta(hours=MAX_SESSION_H):
                    cur['missing'] = True; cur['why'] = 'session longer than 24h'
                cur = None
        if cur is not None and cur['end'] is None:
            cur['missing'] = True; cur['why'] = 'no clock-out'

        # ---- days ----
        byday = collections.OrderedDict()
        for s in sessions:
            byday.setdefault(s['start'].date(), []).append(s)
        for d, ss in byday.items():
            acts = []
            for s in ss:
                for a in s['acts']:
                    if not acts or acts[-1] != a:
                        acts.append(a)
            valid = [x for x in ss if not x['missing'] and x['end'] is not None]
            last = ss[-1]
            out = valid[-1]['end'] if valid else None     # last valid clock-out of the day
            rec = dict(**{'in': ss[0]['start'].isoformat(sep=' ')},
                       out=out.isoformat(sep=' ') if out else None,
                       lok='> '.join(acts) if acts else None)
            days[f'{code}|{d.isoformat()}'] = rec
            if out is None:
                issues['missing_out'].append(dict(code=code, name=last_name, date=d.isoformat(), why=last['why']))
            else:
                hrs = (out - ss[0]['start']).total_seconds() / 3600
                if hrs > 14:
                    issues['long_shifts'].append(dict(code=code, name=last_name, date=d.isoformat(), hours=round(hrs, 1)))
        for x in es:
            if x['backdated']:
                issues['backdated'].append(dict(
                    code=code, name=last_name, type=x['type'], tanggal=x['tanggal'].isoformat(),
                    waktu=x['waktu'].strftime('%Y-%m-%d %H:%M') if x['waktu'] else '',
                    created=x['local'].strftime('%Y-%m-%d %H:%M'), by=x['by']))
    for k in issues:
        issues[k].sort(key=lambda r: (r.get('name', ''), r.get('date', r.get('tanggal', ''))))
    return dict(people=people, days=days, issues=issues)


def parse(data, alias=None, use_waktu=False):
    entries, tanggal = read_entries(data)
    res = build_days(entries, alias, use_waktu)
    res['start'] = tanggal[0].isoformat()
    res['end'] = tanggal[-1].isoformat()
    res['raw_codes'] = sorted({e['code'] for e in entries})
    return res
