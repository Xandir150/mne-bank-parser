#!/usr/bin/env python3
"""Применяет к mapping.json дефолтные разнесения (dio)-счетов, чтобы движки
не задваивали суммы. Правило: счёт остаётся в ОДНОЙ строке группы (см. mapping_notes.md §6),
в остальных строках убирается из include (но остаётся в dio — для флага «проверь руками»).
Запуск: python3 apply_defaults.py   (идемпотентен, правит mapping.json на месте)
"""
import json, os

P = os.path.join(os.path.dirname(__file__), 'mapping.json')
m = json.load(open(P))
S = m['sections']

def row(sec, rb):
    for r in S[sec]['rows']:
        if r['redniBroj'] == rb:
            return r
    raise KeyError(f'{sec}/{rb}')

def drop_include(sec, rb, codes):
    r = row(sec, rb)
    inc = r['accounts']['include']
    r['accounts']['include'] = [c for c in inc if c not in codes]

def add_exclude(sec, rb, codes):
    r = row(sec, rb)
    exc = r['accounts'].setdefault('exclude', [])
    for c in codes:
        if c not in exc:
            exc.append(c)

def add_include(sec, rb, codes):
    r = row(sec, rb)
    inc = r['accounts']['include']
    for c in codes:
        if c not in inc:
            inc.append(c)

BS, BU, SA = 'BilanStanja', 'BilanUspjeha', 'StatistickiAneks'

# --- BilanUspjeha ---
add_exclude(BU, '210', ['540'])                    # 540 живёт в 210a
drop_include(BU, '218', ['589'])                   # весь 589 -> 219
for rb in ('227', '231'): drop_include(BU, rb, ['660'])   # весь 660 -> 223
for rb in ('228', '232'): drop_include(BU, rb, ['661'])   # весь 661 -> 224
for rb in ('229', '233'): drop_include(BU, rb, ['669'])   # весь 669 -> 225
drop_include(BU, '233', ['662', '663', '664'])            # 662-664 -> 229
# 52: 520,521 -> 212 (нето зараде); остальное -> 214 (порези/доприноси в сумме); 215/216 = 0
row(BU, '212')['accounts']['include'] = ['520', '521']
row(BU, '214')['accounts']['include'] = ['522', '523', '524', '525', '526', '529']
row(BU, '215')['accounts']['include'] = []
row(BU, '216')['accounts']['include'] = []
# пробелы формы: 565 (расходы от учешћа, метод удјела) -> 238; 665 (приходи od učešća) -> 223
add_include(BU, '238', ['565'])
add_include(BU, '223', ['665'])

# --- BilanStanja ---
for rb in ('018', '019', '020', '022', '023'): drop_include(BS, rb, ['039'])  # весь 039 -> 017
drop_include(BS, '020', ['033'])                   # весь 033 -> 018
for rb in ('021', '022'): drop_include(BS, rb, ['031', '032'])  # весь 031/032 -> 019
for rb in ('033', '034'): drop_include(BS, rb, ['209'])         # весь 209 -> 032
drop_include(BS, '108', ['322'])                   # весь 322 -> 107
drop_include(BS, '131', ['422', '423', '424', '425', '429'])    # займы -> 130 (небанковские типичнее)
drop_include(BS, '134', ['439'])                   # весь 439 -> 138 (мјенице редки)
drop_include(BS, '126', ['495'])                   # весь 495 -> 143 (краткосрочные ПВР)

# --- StatistickiAneks: 026/027/028 все = 652 -> оставить только 026 ---
for rb in ('027', '028'):
    row(SA, rb)['accounts']['include'] = []

m['_defaults_applied'] = ('dio-splits v1: см. mapping_notes.md §6 и apply_defaults.py; '
                          'разнесение уточняется вручную через Корректировки в обработке')
json.dump(m, open(P, 'w'), ensure_ascii=False, indent=1)
print('OK, defaults applied')
