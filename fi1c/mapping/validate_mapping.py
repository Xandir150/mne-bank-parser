# -*- coding: utf-8 -*-
import json, re, collections, os

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
MAP = os.path.join(BASE, 'fi1c', 'mapping', 'mapping.json')
CHART = os.path.join(BASE, 'fi1c', 'reference', 'mne_chart.tsv')

ACCOUNTS = {}
with open(CHART, encoding='utf-8') as fh:
    for line in fh:
        p = line.rstrip('\n').split('\t')
        if len(p) < 5 or p[0] != 'ACC':
            continue
        ACCOUNTS[p[1]] = (p[2], p[3] == 'True', p[4])
CODES = sorted(ACCOUNTS)

m = json.load(open(MAP, encoding='utf-8'))
secs = m['sections']

report = collections.OrderedDict()

# 1. counts
counts = collections.OrderedDict()
for name, sec in secs.items():
    c = collections.Counter(r['kind'] for r in sec['rows'])
    counts[name] = dict(c)
report['counts'] = counts

# 2. prefixes
def prefix_state(p):
    if p in ACCOUNTS:
        return 'exact'
    if any(c.startswith(p) for c in CODES):
        return 'parent'
    return 'MISSING'

bad_prefix = []
for name, sec in secs.items():
    for r in sec['rows']:
        acc = r.get('accounts')
        if not acc:
            continue
        for field in ('include', 'exclude', 'dio'):
            for p in acc.get(field, []):
                if prefix_state(p) == 'MISSING':
                    bad_prefix.append((name, r['redniBroj'], field, p))
report['missing_prefixes'] = bad_prefix

# 3. formula refs
TOKEN = re.compile(r'[0-9A-Za-z]+')
bad_ref = []
for name, sec in secs.items():
    ids = {r['redniBroj'] for r in sec['rows'] if r['redniBroj']}
    for r in sec['rows']:
        f = r.get('formula')
        if not f:
            continue
        for t in TOKEN.findall(f):
            if t not in ids:
                bad_ref.append((name, r['redniBroj'], f, t))
        if r['kind'] != 'formula':
            bad_ref.append((name, r['redniBroj'], f, 'formula on kind=' + r['kind']))
report['bad_formula_refs'] = bad_ref

# 3b. schema sanity
schema_err = []
VALID_KIND = {'account', 'formula', 'manual', 'header'}
VALID_BAL = {'debitNet', 'creditNet', 'debitOnly', 'creditOnly',
             'turnoverDebit', 'turnoverCredit', 'turnoverCreditNet',
             'turnoverDebitNet'}
for name, sec in secs.items():
    seen = set()
    for r in sec['rows']:
        if r['kind'] not in VALID_KIND:
            schema_err.append((name, r['redniBroj'], 'bad kind ' + r['kind']))
        if r['kind'] == 'account':
            if 'accounts' not in r or not r['accounts'].get('include'):
                schema_err.append((name, r['redniBroj'], 'account w/o include'))
            if r.get('balance') not in VALID_BAL:
                schema_err.append((name, r['redniBroj'], 'bad balance'))
        else:
            if 'accounts' in r:
                schema_err.append((name, r['redniBroj'], 'accounts on non-account'))
            if 'balance' in r:
                schema_err.append((name, r['redniBroj'], 'balance on non-account'))
        if r['kind'] == 'formula' and not r.get('formula'):
            schema_err.append((name, r['redniBroj'], 'formula kind w/o formula'))
        rb = r['redniBroj']
        if rb:
            if rb in seen:
                schema_err.append((name, rb, 'duplicate redniBroj'))
            seen.add(rb)
report['schema_errors'] = schema_err

# 4. coverage
def row_covers(a, acc):
    inc = acc['include']
    exc = acc.get('exclude', [])
    inside = any(a.startswith(p) for p in inc)
    if inside and any(a.startswith(e) for e in exc):
        inside = False
    parent = any(p.startswith(a) for p in inc)
    return inside or parent

def acc_rows(name):
    return [r for r in secs[name]['rows'] if r['kind'] == 'account']

BS = acc_rows('BilanStanja')
BU = acc_rows('BilanUspjeha')
SA = acc_rows('StatistickiAneks')

uncovered = []
cover_index = collections.defaultdict(list)
for code, (typ, off, nm) in sorted(ACCOUNTS.items()):
    if off or code[0] not in '0123456':
        continue
    pool = BS if code[0] in '01234' else BU
    hits = [r['redniBroj'] for r in pool if row_covers(code, r['accounts'])]
    if hits:
        cover_index[code] = hits
    else:
        alt = [r['redniBroj'] for r in SA if row_covers(code, r['accounts'])]
        uncovered.append((code, nm[:70], 'StatAneks:' + ','.join(alt) if alt else ''))
report['uncovered'] = uncovered

# 5. dio rows
dio_rows = []
for name, sec in secs.items():
    for r in sec['rows']:
        acc = r.get('accounts')
        if acc and acc.get('dio'):
            dio_rows.append((name, r['redniBroj'], r['pozicija'][:60],
                             r['grupaRacuna'], ','.join(acc['dio'])))
report['dio_rows'] = dio_rows

# 6. accounts hit by more than one account-row in the same section
multi = []
for name in ('BilanStanja', 'BilanUspjeha'):
    pool = acc_rows(name)
    for code, (typ, off, nm) in sorted(ACCOUNTS.items()):
        if off:
            continue
        hits = [r for r in pool if code in r['accounts']['include']]
        if len(hits) > 1:
            multi.append((name, code, ','.join(h['redniBroj'] for h in hits)))
report['multi_mapped'] = multi

for k, v in report.items():
    print('===', k)
    if isinstance(v, dict):
        for kk, vv in v.items():
            print('   ', kk, vv)
    else:
        print('    n =', len(v))
        for item in v:
            print('   ', item)
