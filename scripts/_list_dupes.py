import openpyxl
from collections import Counter

wb = openpyxl.load_workbook(r"C:\Feven\Hanna\2026\Gainwell Servers_20260820.xlsx", read_only=True, data_only=True)
ws = wb["data"]
rows = ws.iter_rows(values_only=True)
header = list(next(rows))
name_idx = header.index("Name")

names = []
for row in rows:
    v = row[name_idx]
    if v is not None and str(v).strip() != "":
        names.append(str(v).strip())

counts = Counter(n.upper() for n in names)
dupe_upper = {n for n, c in counts.items() if c > 1}

original_casing = {}
for n in names:
    original_casing.setdefault(n.upper(), set()).add(n)

print(f"Duplicate ServerName values in source Excel: {len(dupe_upper)}")
for n in sorted(dupe_upper):
    variants = original_casing[n]
    print(f"  {n}  (as written: {sorted(variants)}, {counts[n]} rows)")
