# -*- coding: utf-8 -*-
"""Generate fi1c/mapping/mapping.json from xsd/fi.xml + reference/mne_chart.tsv."""
import json, re, os, collections
import xml.etree.ElementTree as ET

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
FI = os.path.join(BASE, 'xsd', 'fi.xml')
CHART = os.path.join(BASE, 'fi1c', 'reference', 'mne_chart.tsv')
OUT = os.path.join(BASE, 'fi1c', 'mapping', 'mapping.json')

# ---------------------------------------------------------------- chart ----
ACCOUNTS = {}          # code -> (type, offbalance, name)
with open(CHART, encoding='utf-8') as fh:
    for line in fh:
        p = line.rstrip('\n').split('\t')
        if len(p) < 5 or p[0] != 'ACC':
            continue
        ACCOUNTS[p[1]] = (p[2], p[3] == 'True', p[4])
CODES = sorted(ACCOUNTS)
ONBAL = [c for c in CODES if not ACCOUNTS[c][1]]

def exists(prefix):
    """True if prefix is a real account or a parent/child of one."""
    return any(c == prefix or c.startswith(prefix) for c in CODES)

# ------------------------------------------------------- GrupaRacuna parse --
DIO_RE = re.compile(r'^(\d+)\s*\(\s*dio\s*\)$', re.I)
RANGE_RE = re.compile(r'^(\d+)\s*-\s*(\d+)$')
CODE_RE = re.compile(r'^\d+$')
SPLIT_EXCL = re.compile(r'\b(?:osim|sem|bez)\b', re.I)

def expand_range(lo, hi):
    """440-449 -> only codes that really exist at that digit length."""
    n = len(lo)
    out = [str(v).zfill(n) for v in range(int(lo), int(hi) + 1)]
    real = [c for c in out if c in ACCOUNTS]
    return real if real else out

def parse_tokens(text):
    """-> (codes, dio_codes) for one comma/'i'-separated fragment list."""
    codes, dio = [], []
    text = text.replace('и', ' i ')                      # cyrillic 'и'
    parts = re.split(r',|\bi\b', text, flags=re.I)
    for raw in parts:
        tok = raw.strip().strip('.').strip()
        if not tok:
            continue
        m = DIO_RE.match(tok)
        if m:
            codes.append(m.group(1)); dio.append(m.group(1)); continue
        m = RANGE_RE.match(tok)
        if m:
            codes.extend(expand_range(m.group(1), m.group(2))); continue
        if CODE_RE.match(tok):
            codes.append(tok); continue
        raise ValueError('unparsed token %r in %r' % (tok, text))
    return codes, dio

def parse_grupa(text):
    """GrupaRacuna string -> accounts dict."""
    head, *tail = SPLIT_EXCL.split(text)
    inc, dio = parse_tokens(head)
    exc = []
    for chunk in tail:
        c, d = parse_tokens(chunk)
        exc.extend(c); dio.extend(d)
    acc = {'include': dedup(inc)}
    if exc:
        acc['exclude'] = dedup(exc)
    if dio:
        acc['dio'] = dedup(dio)
    return acc

def dedup(seq):
    seen, out = set(), []
    for x in seq:
        if x not in seen:
            seen.add(x); out.append(x)
    return out

# ------------------------------------------------------------- rule tables --
BS_FORMULA = {
    '002': '003+008+016',
    '003': '004+005+006+007',
    '008': '009+010+011+015',
    '011': '012+013+014',
    '016': '017+018+019+020+021+022+023',
    '025': '026+031+039+043+044',
    '026': '027+028+029+030',
    '031': '032+033+034+035',
    '035': '036+037+038',
    '039': '040+041+042',
    '046': '001+002+024+025+045',
    '101': '102+103+104+105+111+116',
    '105': '106+107+108+109-110',
    '111': '112+113-114-115',
    '117': '118+122',
    '118': '119+120+121',
    '122': '123+124',
    '127': '128+129',
    '129': '130+131+132+133+134+135+136+137',
    '137': '138+139+140+141+142',
    '144': '101+117+125+126+127+143',
}
BS_MANUAL = {'116'}
# explicit account overrides (parser cannot read the cyrillic saldo phrases)
BS_ACC_OVERRIDE = {
    '109': {'include': ['330', '331', '332', '333', '334', '335', '336', '337']},
    '110': {'include': ['331', '332', '333', '334', '335', '336', '337']},
}
BS_BALANCE_OVERRIDE = {
    '109': 'creditOnly',
    '110': 'debitOnly',
    # 350/351 carry debit balances and are subtracted by formula 111 -> must be
    # reported as positive debit amounts (documented deviation, see notes).
    '114': 'debitNet',
    '115': 'debitNet',
}

BU_FORMULA = {
    '204': '205+206+207',
    '208': '209+210+210a',
    '211': '212+213',
    '213': '214+215+216',
    '217': '218+219',
    '221': '201+202+203+204-208-211-217-220',
    '222': '223+224+225',
    '226': '227+228+229',
    '230': '231+232+233',
    '234': '235-236',
    '237': '238+239+240',
    '241': '222+226+230+234-237',
    '242': '221+241',
    '244': '242+243',
    '245': '246+247',
    '248': '244-245',
    '249': '250+251+252+253+254+255+256+257',
    '259': '249+258',
    '260': '248-259',
}
BU_MANUAL = {'258', '261', '262', '263', '264', '265'}
BU_ACC_OVERRIDE = {
    # "690-590": credit-net over both accounts == income(690) - expense(590)
    '243': {'include': ['690', '590']},
}
BU_BALANCE = {
    '201': 'turnoverCreditNet', '202': 'turnoverCreditNet',
    '203': 'turnoverCreditNet', '205': 'turnoverCreditNet',
    '206': 'turnoverCreditNet', '207': 'turnoverCreditNet',
    '209': 'turnoverDebitNet', '210': 'turnoverDebitNet',
    '210a': 'turnoverDebitNet', '212': 'turnoverDebitNet',
    '214': 'turnoverDebitNet', '215': 'turnoverDebitNet',
    '216': 'turnoverDebitNet', '218': 'turnoverDebitNet',
    '219': 'turnoverDebitNet', '220': 'turnoverDebitNet',
    '223': 'turnoverCreditNet', '224': 'turnoverCreditNet',
    '225': 'turnoverCreditNet', '227': 'turnoverCreditNet',
    '228': 'turnoverCreditNet', '229': 'turnoverCreditNet',
    '231': 'turnoverCreditNet', '232': 'turnoverCreditNet',
    '233': 'turnoverCreditNet', '235': 'turnoverCreditNet',
    '236': 'turnoverDebitNet', '238': 'turnoverDebitNet',
    '239': 'turnoverDebitNet', '240': 'turnoverDebitNet',
    '243': 'turnoverCreditNet',
    '246': 'turnoverDebitNet', '247': 'turnoverDebitNet',
    '250': 'turnoverCreditNet', '251': 'turnoverCreditNet',
    '252': 'turnoverCreditNet', '253': 'turnoverCreditNet',
    '254': 'turnoverCreditNet', '255': 'turnoverCreditNet',
    '256': 'turnoverCreditNet', '257': 'turnoverCreditNet',
}

SA_CONTRA = {'034', '035', '042', '043', '047', '053', '054', '059'}
SA_REVENUE = {'002', '003', '004', '005', '006', '007', '026', '027', '028'}
SA_EXPENSE = {'008', '009', '010', '011', '012', '013', '014', '015',
              '016', '017', '018', '019', '020', '021'}

CF_FORMULA = {
    '301': '302+303+304',
    '305': '306+307+308+309+310',
    '311': '301-305',
    '312': '313+314+315+316+317',
    '318': '319+320+321',
    '322': '312-318',
    '323': '324+325+326',
    '327': '328+329+330+331',
    '332': '323-327',
    '333': '311+322+332',
    '337': '333+334+335-336',
}
CF3A_FORMULA = dict(CF_FORMULA)
CF3A_FORMULA['311'] = '301+302+303+304+305+306+307+308+309+310'
del CF3A_FORMULA['301']
del CF3A_FORMULA['305']

KAP_FORMULA = {'3': '1+2', '5': '3+4', '7': '5+6', '9': '7+8'}

# ------------------------------------------------------------------ build --
tree = ET.parse(FI)
root = tree.getroot()

def txt(node, tag):
    v = node.findtext(tag)
    return (v or '').strip()

sections = collections.OrderedDict()
problems = []

def build_plain(secname, formula_tbl, manual_set, acc_override,
                balance_fn, balance_override=None):
    rows = []
    for st in root.find(secname):
        rb = txt(st, 'RedniBroj')
        poz = txt(st, 'Pozicija')
        gr = txt(st, 'GrupaRacuna')
        row = {'redniBroj': rb, 'pozicija': poz, 'grupaRacuna': gr}
        if not rb:
            row['kind'] = 'header'
        elif rb in formula_tbl:
            row['kind'] = 'formula'
            row['formula'] = formula_tbl[rb]
        elif rb in acc_override:
            row['kind'] = 'account'
            row['accounts'] = acc_override[rb]
        elif rb in manual_set or not gr:
            row['kind'] = 'manual'
        else:
            row['kind'] = 'account'
            try:
                row['accounts'] = parse_grupa(gr)
            except ValueError as e:
                problems.append('%s %s: %s' % (secname, rb, e))
                row['kind'] = 'manual'
        if row['kind'] == 'account':
            b = (balance_override or {}).get(rb) or balance_fn(rb, poz, gr)
            row['balance'] = b
        rows.append(row)
    return rows

# --- BilanStanja
def bs_balance(rb, poz, gr):
    return 'debitNet' if rb < '100' else 'creditNet'

sections['BilanStanja'] = {'rows': build_plain(
    'BilanStanja', BS_FORMULA, BS_MANUAL, BS_ACC_OVERRIDE,
    bs_balance, BS_BALANCE_OVERRIDE)}

# --- BilanUspjeha
def bu_balance(rb, poz, gr):
    b = BU_BALANCE.get(rb)
    if not b:
        problems.append('BilanUspjeha %s: no balance rule, defaulted' % rb)
        b = 'turnoverCreditNet'
    return b

sections['BilanUspjeha'] = {'rows': build_plain(
    'BilanUspjeha', BU_FORMULA, BU_MANUAL, BU_ACC_OVERRIDE, bu_balance)}

# --- cash-flow statements (no GrupaRacuna at all)
def build_cf(secname, formula_tbl):
    rows = []
    for st in root.find(secname):
        rb = txt(st, 'RedniBroj')
        poz = txt(st, 'Pozicija')
        row = {'redniBroj': rb, 'pozicija': poz, 'grupaRacuna': ''}
        if not rb:
            row['kind'] = 'header'
        elif rb in formula_tbl:
            row['kind'] = 'formula'
            row['formula'] = formula_tbl[rb]
        else:
            row['kind'] = 'manual'
        rows.append(row)
    return rows

sections['IskazOTokovimaGotovine'] = {
    'rows': build_cf('IskazOTokovimaGotovine', CF_FORMULA)}
sections['TokoviGotovine3a'] = {
    'rows': build_cf('TokoviGotovine3a', CF3A_FORMULA)}

# --- IskazOPromenamaKapitala (matrix: Pozicija = row no, Opis = label)
kap_rows = []
for st in root.find('IskazOPromenamaKapitala'):
    rb = txt(st, 'Pozicija').rstrip('.')
    row = {'redniBroj': rb, 'pozicija': txt(st, 'Opis'), 'grupaRacuna': ''}
    if rb in KAP_FORMULA:
        row['kind'] = 'formula'
        row['formula'] = KAP_FORMULA[rb]
    else:
        row['kind'] = 'manual'
    kap_rows.append(row)
sections['IskazOPromenamaKapitala'] = {'rows': kap_rows}

# --- StatistickiAneks
def sa_balance(rb, poz, gr):
    if rb in SA_REVENUE:
        return 'turnoverCreditNet'
    if rb in SA_EXPENSE:
        return 'turnoverDebitNet'
    if rb in SA_CONTRA:
        return 'creditNet'
    return 'debitNet'

sections['StatistickiAneks'] = {'rows': build_plain(
    'StatistickiAneks', {}, set(), {}, sa_balance)}

# --- ObracunAmortizacije (Grupa I..V, no positions/accounts in the template)
am_rows = []
for st in root.find('ObracunAmortizacije'):
    g = txt(st, 'Grupa')
    stopa = txt(st, 'Iznos6')
    am_rows.append({
        'redniBroj': g,
        'pozicija': 'Amortizaciona grupa %s (stopa %s%%)' % (g, stopa),
        'grupaRacuna': '',
        'kind': 'manual',
    })
sections['ObracunAmortizacije'] = {'rows': am_rows}

mapping = {'sections': sections}
os.makedirs(os.path.dirname(OUT), exist_ok=True)
with open(OUT, 'w', encoding='utf-8') as fh:
    json.dump(mapping, fh, ensure_ascii=False, indent=2)
    fh.write('\n')

print('written', OUT)
for p in problems:
    print('PROBLEM:', p)
