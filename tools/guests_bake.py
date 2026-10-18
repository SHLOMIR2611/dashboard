#!/usr/bin/env python3
"""v331: bake the Datwise "הדרכת אורחים" data into index.html as const GUESTS.
Usage: guests_bake.py <index.html> <qualifications xlsx exported with NO status filter, 'כל המפגשים שבוצעו'> "DD/MM/YYYY HH:MM"
Splits people who did the guest training by their current Datwise status:
  אורח/ת  -> real guests (g)
  פעיל/ה  -> candidates who were hired (h)
  other   -> no longer active (x)
No ID numbers or phone numbers are written to the page."""
import sys, re, json, datetime, openpyxl
html, xlsx, asof = sys.argv[1], sys.argv[2], sys.argv[3]
ws = openpyxl.load_workbook(xlsx, read_only=True).active
rows = list(ws.iter_rows(values_only=True))
H = list(rows[0]); ix = {k: H.index(k) for k in H if k}
def c(r, k): v = r[ix[k]]; return '' if v is None else str(v).strip()
def d(s):
    try: return datetime.datetime.strptime(s, '%d/%m/%Y')
    except Exception: return None
best = {}
for r in rows[1:]:
    if c(r, 'נושא') != 'הדרכת אורחים': continue
    key = c(r, 'מספר ת.ז. / דרכון') or (c(r, 'שם פרטי') + ' ' + c(r, 'שם משפחה'))
    if key not in best or (d(c(r, 'תאריך')) or datetime.datetime.min) > (d(c(best[key], 'תאריך')) or datetime.datetime.min):
        best[key] = r
G = {'asof': asof, 'g': [], 'h': [], 'x': []}
for r in best.values():
    name = (c(r, 'שם פרטי') + ' ' + c(r, 'שם משפחה')).strip()
    base = [name, c(r, 'חברת שירות / קבלן'), c(r, 'תאריך'), c(r, 'תאריך פקיעת תוקף'), c(r, 'סטטוס משתתף')]
    st = c(r, 'סטטוס')
    if st == 'אורח/ת': G['g'].append(base)
    elif st == 'פעיל/ה': G['h'].append(base + [c(r, 'חטיבה'), c(r, 'מחלקה'), c(r, 'מנהל/ת'), c(r, 'תאריך תחילת עבודה')])
    else: G['x'].append(base + [st])
for k in ('g', 'h', 'x'):
    G[k].sort(key=lambda a: d(a[2]) or datetime.datetime.min, reverse=True)
s = open(html, encoding='utf-8').read()
line = 'const GUESTS = ' + json.dumps(G, ensure_ascii=False, separators=(',', ':')) + ';'
if re.search(r'^const GUESTS = .*;$', s, re.M):
    s = re.sub(r'^const GUESTS = .*;$', lambda m: line, s, count=1, flags=re.M)
else:
    i = s.index('const ST = '); s = s[:i] + line + '\n' + s[i:]
open(html, 'w', encoding='utf-8').write(s)
print(json.dumps({'guests': len(G['g']), 'hired': len(G['h']), 'inactive': len(G['x'])}))
