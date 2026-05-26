import glob, traceback, re
from pathlib import Path
from app.parsers.nlb import NLBParser, _build_cid_maps_from_font
import pdfplumber

f = sorted(glob.glob("/data/input/*64*.pdf"))[0]
print(f"File: {f}")

# Check maps
b, r = _build_cid_maps_from_font(f)
print(f"Bold map: {len(b)} entries")
print(f"Regular map: {len(r)} entries")
print(f"Bold: {dict(sorted(b.items())[:10])}")
print(f"Regular: {dict(sorted(r.items())[:10])}")

# Check tables
with pdfplumber.open(f) as pdf:
    page = pdf.pages[0]
    tables = page.find_tables()
    print(f"\nTables found: {len(tables)}")
    for i, table in enumerate(tables):
        cells = table.cells
        xs = sorted(set(c[0] for c in cells))
        ys = sorted(set(c[1] for c in cells))
        print(f"  Table {i}: {len(xs)} cols x {len(ys)} rows, cells={len(cells)}")

    ext_tables = page.extract_tables()
    print(f"\nextract_tables: {len(ext_tables)}")
    for i, t in enumerate(ext_tables):
        print(f"  Table {i}: {len(t)} rows, cols={len(t[0]) if t else 0}")
        if t:
            # Decode first cell
            cell = t[0][0] or ""
            decoded = re.sub(r"\(cid:(\d+)\)", lambda m: r.get(int(m.group(1)), b.get(int(m.group(1)), "?")), cell)
            print(f"    First cell decoded: {decoded[:80]}")

# Full parse
p = NLBParser()
s = p.parse(Path(f))
print(f"\nParsed: txns={len(s.transactions)} acct={s.account_number} stmt={s.statement_number}")
