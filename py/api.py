"""Entry points used by the web page (via Pyodide) and by tools/cli.py.

Everything goes through JSON strings so the JavaScript side stays simple.

  load(csv_path, prev_path=None, use_waktu=False) -> analysis JSON
  merge(keep_code, absorb_json)                   -> analysis JSON (after applying the merge)
  convert(template_path, out_path, edits_json)    -> summary JSON
"""
import io, json, datetime as dt
import openpyxl
import parse, roster, build, config

STATE = {}


def _load_prev(prev_path):
    if not prev_path:
        return [], [], 'none'
    with open(prev_path, 'rb') as f:
        wb = openpyxl.load_workbook(io.BytesIO(f.read()))
    return roster.read_roster(wb)


def _analyze():
    st = STATE
    parsed = parse.build_days(st['entries'], roster.alias_map(st['roster']), st['use_waktu'])
    parsed['start'] = st['tanggal'][0].isoformat(); parsed['end'] = st['tanggal'][-1].isoformat()
    groups = st['groups'] or config.GROUP_ORDER       # no roster: built-in project order
    rec = roster.reconcile(st['roster'], groups, parsed)
    for p in rec['people']:
        if p['code'] in config.PIN_FIRST and p['order'] is None:
            p['order'] = -1
    st['parsed'] = parsed; st['rec'] = rec
    people = roster.sort_people(rec['people'], rec['groups'])
    start = dt.date.fromisoformat(parsed['start']); end = dt.date.fromisoformat(parsed['end'])
    return json.dumps(dict(
        period=dict(start=parsed['start'], end=parsed['end'], ndays=(end - start).days + 1),
        roster_source=st['roster_source'], roster_size=len(st['roster']),
        people=[{k: p[k] for k in ('code', 'name', 'grup', 'grups', 'codes', 'jabatan', 'sk', 'new',
                                   'moved_from', 'multi_group', 'matched_by', 'rid')} for p in people],
        groups=rec['groups'], new_groups=rec['new_groups'],
        not_in_export=rec['not_in_export'], duplicates=rec['duplicates'],
        issues=parsed['issues'], n_entries=len(st['entries'])), default=str)


def _saved(saved_json):
    """Roster remembered by the browser from the last conversion (same shape write_roster uses)."""
    try:
        d = json.loads(saved_json) if saved_json else None
    except ValueError:
        d = None
    if not d or not d.get('entries'):
        return [], [], 'none'
    keys = ('code', 'nama', 'grup', 'jabatan', 'sk', 'first_seen', 'last_seen')
    ents = [dict({k: str(e.get(k) or '') for k in keys}, active=bool(e.get('active', True)),
                 order=e.get('order'), alias=list(e.get('alias') or [])) for e in d['entries']]
    return ents, list(d.get('groups') or []), 'saved'


def load(csv_path, prev_path=None, use_waktu=False, saved_roster_json=None):
    """Roster source: last month's workbook if given, else the roster the browser remembered."""
    with open(csv_path, 'rb') as f:
        entries, tanggal = parse.read_entries(f.read())
    entries_roster, groups, source = _load_prev(prev_path)
    if source == 'none' and not prev_path:
        entries_roster, groups, source = _saved(saved_roster_json)
    STATE.clear()
    STATE.update(entries=entries, tanggal=tanggal, roster=entries_roster, groups=groups,
                 roster_source=source, use_waktu=bool(use_waktu))
    return _analyze()


def set_option(use_waktu):
    STATE['use_waktu'] = bool(use_waktu)
    return _analyze()


def merge(keep_code, absorb_json):
    """Treat two Kode Anggota as the same person. keep_code is this month's code.
    absorb: {"kind": "roster", "rid": n} or {"kind": "csv", "code": "X-1"}."""
    absorb = json.loads(absorb_json)
    ents = STATE['roster']
    keep_i = next((i for i, e in enumerate(ents) if e['code'] == keep_code), None)
    if absorb['kind'] == 'roster':
        a = ents[int(absorb['rid'])]
        merged = dict(a)
        merged['alias'] = list(a.get('alias', []))
        if a['code'] and a['code'] != keep_code:
            merged['alias'].append(a['code'])
        merged['code'] = keep_code
        merged['active'] = True
        drop = {int(absorb['rid'])}
        if keep_i is not None and keep_i != int(absorb['rid']):
            merged['alias'] += [x for x in ents[keep_i].get('alias', []) if x not in merged['alias']]
            drop.add(keep_i)
        STATE['roster'] = [e for i, e in enumerate(ents) if i not in drop] + [merged]
    else:
        other = absorb['code']
        if keep_i is None:
            p = next(p for p in STATE['parsed']['people'] if p['code'] == keep_code)
            ents.append(dict(code=keep_code, nama=p['name'].upper(), grup=p['grup'], jabatan='', sk='',
                             active=True, first_seen='', last_seen='', order=None, alias=[]))
            keep_i = len(ents) - 1
        if other not in ents[keep_i]['alias']:
            ents[keep_i]['alias'].append(other)
        # if the absorbed code had its own roster entry, fold it in
        oi = next((i for i, e in enumerate(ents) if e['code'] == other), None)
        if oi is not None:
            o = ents[oi]
            for k in ('jabatan', 'sk', 'order', 'first_seen'):
                if not ents[keep_i].get(k) and o.get(k):
                    ents[keep_i][k] = o[k]
            ents[keep_i]['alias'] += [x for x in o.get('alias', []) if x not in ents[keep_i]['alias']]
            del ents[oi]
    return _analyze()


def convert(template_path, out_path, edits_json):
    """edits: {people: {code: {jabatan, sk, ket}}, groups: [...], manual: [...], options: {...}}"""
    edits = json.loads(edits_json)
    parsed, rec = STATE['parsed'], STATE['rec']
    groups = edits.get('groups') or rec['groups']
    for g in rec['groups']:
        if g not in groups:
            groups.append(g)
    pe = edits.get('people', {})
    manual = [m for m in edits.get('manual', []) if m.get('code') and m.get('date') and m.get('status')]
    opts = edits.get('options', {})

    # final people list
    people = []
    for p in roster.sort_people(rec['people'], groups):
        e = pe.get(p['code'], {})
        jab = (e.get('jabatan', p['jabatan']) or '').strip()
        sk = (e.get('sk', p['sk']) or '').strip()
        ket_parts = []
        mine = [(dt.date.fromisoformat(m['date']), m['status'].upper(), (m.get('note') or '').strip())
                for m in manual if m['code'] == p['code']]
        if mine:
            ket_parts.append(build.ket_from_manual(mine))
        if (e.get('ket') or '').strip():
            ket_parts.append(e['ket'].strip())
        people.append(dict(code=p['code'], name=p['name'].upper(), grup=p['grup'], jabatan=jab, sk=sk,
                           jab=(jab + '\n' if jab else '') + 'Proyek : ' + p['grup'],
                           ket='\n'.join(ket_parts) or None, rid=p['rid'], order=p['order']))

    with open(template_path, 'rb') as f:
        wb, summary = build.build(f.read(), parsed, people, manual, opts)

    # ---- roster for next month ----
    old = STATE['roster']
    out_entries, used = [], set()
    for i, q in enumerate(people, start=1):
        r = old[q['rid']] if q['rid'] is not None else None
        if r is not None:
            used.add(q['rid'])
        out_entries.append(dict(
            code=q['code'], nama=q['name'], grup=q['grup'], jabatan=q['jabatan'], sk=q['sk'], active=True,
            first_seen=(r or {}).get('first_seen') or parsed['start'], last_seen=parsed['end'],
            order=float(i), alias=list((r or {}).get('alias', [])) +
            [c for c in next((p['codes'] for p in rec['people'] if p['code'] == q['code']), []) if c != q['code']]))
    for i, r in enumerate(old):
        if i in used:
            continue
        out_entries.append(dict(r, active=False))
    for e in out_entries:                               # de-dup alias lists
        e['alias'] = sorted({a for a in e['alias'] if a and a != e['code']})
    roster.write_roster(wb, out_entries, groups)

    data = build.save(wb)
    with open(out_path, 'wb') as f:
        f.write(data)
    summary['size'] = len(data)
    summary['groups'] = groups
    summary['roster'] = dict(entries=out_entries, groups=groups)
    return json.dumps(summary, default=str)
