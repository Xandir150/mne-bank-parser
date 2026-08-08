#!/usr/bin/env python3
"""Сравнение результатов двух независимых движков FI:
  compare.py <rows_epf.json> <rows_ficalc.json> [tolerance]
Оба файла: {"sections":{"BilanStanja":{"001":[i1,i2,i3],...},...}} (лишние ключи игнорируются).
Выход 0 — совпало (в пределах допуска, по умолчанию 0.01), 1 — есть расхождения (печатает их).
"""
import json, sys

a = json.load(open(sys.argv[1]))
b = json.load(open(sys.argv[2]))
tol = float(sys.argv[3]) if len(sys.argv) > 3 else 0.01
sa, sb = a.get('sections', a), b.get('sections', b)

diffs = []
for sec in sorted(set(sa) | set(sb)):
    ra, rb = sa.get(sec, {}), sb.get(sec, {})
    for rn in sorted(set(ra) | set(rb)):
        va, vb = ra.get(rn), rb.get(rn)
        if va is None or vb is None:
            diffs.append(f'{sec}/{rn}: присутствует только в одном ({"epf" if vb is None else "ficalc"})')
            continue
        for i in range(max(len(va), len(vb))):
            x = va[i] if i < len(va) else 0.0
            y = vb[i] if i < len(vb) else 0.0
            if abs(float(x) - float(y)) > tol:
                diffs.append(f'{sec}/{rn} Iznos{i+1}: epf={x} ficalc={y} diff={float(x)-float(y):.2f}')

if diffs:
    print(f'РАСХОЖДЕНИЯ: {len(diffs)}')
    for d in diffs[:80]:
        print(' ', d)
    if len(diffs) > 80:
        print(f'  ... и ещё {len(diffs)-80}')
    sys.exit(1)
print('OK: движки сходятся')
