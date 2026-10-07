"""E[emp]=[name,role,manager,dept,division]. Rebuild TRN_FULL (anti-cookies tab) from Blossom files.
usage: python3 tools/blossom_build.py <index.html> <csv_dir> <basic.xlsx> "DD/MM/YYYY HH:MM"
"""
import csv,glob,re,json,openpyxl,sys,datetime
P,CSVDIR,XLSX,STAMP=sys.argv[1:5]
s=open(P,encoding='utf-8').read()
i=s.index('const TRN_FULL = ')+len('const TRN_FULL = '); old,ln=json.JSONDecoder().raw_decode(s[i:]); end=i+ln
oldcat={t['t'].strip():t['c'] for t in old['T'] if t['c'] in('all','prod','hr')}
KW=['עובדי משרד אינפק','בד"ס','בדיקות ראייה','הנדסה/פיתוח','חומרים מסוכנים','מבקרת QC','מיקרוגל','מלחימות','מעבדה- סיכונים','מעבדה-סיכונים','תהליך מיוחד']
isfront=lambda t: any(k in t for k in KW) or t.strip().startswith('PP')
def fd(v):
    if v is None: return ''
    if isinstance(v,datetime.datetime): return v.strftime('%d/%m/%Y')
    return str(v).strip()[:10]
ST={'אין התקדמות':0,'הצלחה':1,'לא הושלם':2,'לא פעיל':3}
E={};T=[]
for f in sorted(glob.glob(CSVDIR+'/*.csv')):
    rows=list(csv.reader(open(f,encoding='utf-8-sig'))); title=rows[0][0].strip(); h=rows[1]
    ix={k:h.index(k) for k in ['משתמש','משרה/תפקיד','מספר עובד','שם מנהל ישיר','תאור מחלקה','סטטוס']}
    di=h.index('תאריך ביצוע') if 'תאריך ביצוע' in h else None
    r=[]
    for x in rows[2:]:
        if len(x)<=ix['סטטוס'] or x[ix['סטטוס']] not in ST: continue
        e=x[ix['מספר עובד']].strip()
        if not e: continue
        dv=h.index('תאור חטיבה') if 'תאור חטיבה' in h else None
        E.setdefault(e,[x[ix['משתמש']].strip(),x[ix['משרה/תפקיד']].strip(),x[ix['שם מנהל ישיר']].strip(),x[ix['תאור מחלקה']].strip(),(x[dv].strip() if dv is not None and len(x)>dv else '')])
        if dv is not None and len(x)>dv and x[dv].strip() and not E[e][4]: E[e][4]=x[dv].strip()
        st=ST[x[ix['סטטוס']]]
        r.append([e,st,fd(x[di]) if di is not None else '',''] if st==1 else [e,st])
    T.append({'t':title,'c':oldcat.get(title,'all'),'r':r})
DST={'הושלם':1,'לא הושלם':2,'פג תוקף':4,'חידוש':5,'באיחור':6}
rows=list(openpyxl.load_workbook(XLSX,read_only=True).active.iter_rows(values_only=True)); h=list(rows[1])
c=lambda n:h.index(n); mgr=c('שם מנהל ישיר') if 'שם מנהל ישיר' in h else c('מנהל ישיר')
dw={}
for x in rows[2:]:
    if not x[0] or x[c('סטטוס')] not in DST or x[c('מספר עובד')] is None: continue
    v=x[c('מספר עובד')]; e=str(int(v)) if isinstance(v,(int,float)) else str(v).strip()
    E.setdefault(e,[str(x[0]).strip(),(x[c('משרה/תפקיד')] or '').strip(),(x[mgr] or '').strip(),(x[c('תאור מחלקה')] or x[c('תאור חטיבה')] or '').strip(),(x[c('תאור חטיבה')] or '').strip()])
    if len(E[e])<5: E[e].append('')
    if not E[e][4] and x[c('תאור חטיבה')]: E[e][4]=str(x[c('תאור חטיבה')]).strip()
    st=DST[x[c('סטטוס')]]
    dw.setdefault(str(x[c('הכשרה')]).strip(),[]).append([e,1,fd(x[c('תאריך השלמה')]),fd(x[c('תאריך יעד')])] if st==1 else [e,st])
for t,r in dw.items(): T.append({'t':t,'c':'dw' if isfront(t) else 'bq','r':r})
s=s[:i]+json.dumps({'E':E,'T':T},ensure_ascii=False,separators=(',',':'))+s[end:]
s,n=re.subn(r'const TRN_FULL_SOURCE = "[^\n]*";','const TRN_FULL_SOURCE = "הורדה ישירה מבלוסום ('+STAMP+'): 18 דוחות \\"דוח - מבצעים\\" + \\"הכשרות שנתיות- לקלוד\\" (דוח הכשרה בסיסי)";',s,count=1); assert n==1
open(P,'w',encoding='utf-8').write(s)
from collections import Counter
print('T',len(T),Counter(t['c'] for t in T),'E',len(E),'done w/ date',sum(1 for t in T for x in t['r'] if x[1]==1 and x[2]))
