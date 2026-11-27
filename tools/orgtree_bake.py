"""B499: rebuild ORGTREE_FULL (מבנה ארגוני) from the native Datwise HR export.
usage: orgtree_bake.py <index.html> <_dw_hr.xlsx>
Active employees only (סטטוס פעיל/ה); tree = אגף > חטיבה > מחלקה > תת מחלקה > [employees]."""
import sys,re,json,openpyxl
html,hr=sys.argv[1],sys.argv[2]
ws=openpyxl.load_workbook(hr,read_only=True).active
rows=list(ws.iter_rows(values_only=True)); H=[str(c or '').strip() for c in rows[0]]
ix={h:i for i,h in enumerate(H)}
def g(r,k):
    v=r[ix[k]] if k in ix and ix[k]<len(r) else ''
    return str(v).strip() if v is not None else ''
T={}; n=0
for r in rows[1:]:
    if g(r,'סטטוס')!='פעיל/ה': continue
    eid=g(r,'מספר עובד/ת'); eid=eid.lstrip('0') or eid
    e={'name':(g(r,'שם פרטי')+' '+g(r,'שם משפחה')).strip(),'eid':eid,'role':g(r,'תפקיד ראשי'),
       'attendance':g(r,'נוכחות'),'manager':g(r,'מנהל/ת')}
    ag=g(r,'אגף') or 'ללא אגף'
    T.setdefault(ag,{}).setdefault(g(r,'חטיבה'),{}).setdefault(g(r,'מחלקה'),{}).setdefault(g(r,'תת מחלקה'),[]).append(e); n+=1
s=open(html,encoding='utf-8').read()
s,k=re.subn(r'const ORGTREE_FULL = \{.*?\};\n',lambda m:'const ORGTREE_FULL = '+json.dumps(T,ensure_ascii=False,separators=(',',':'))+';\n',s,count=1,flags=re.S)
assert k==1
open(html,'w',encoding='utf-8').write(s)
print(json.dumps({'employees':n,'agafim':len(T),'present':sum(1 for a in T.values() for d in a.values() for p in d.values() for L in p.values() for e in L if e['attendance']=='נוכח')},ensure_ascii=False))
