"""Bake native Datwise exports into index.html (via the page's own upload parsers).
usage: python3 tools/dw_bake.py <index.html> <hr.xlsx> <comp.xlsx> <events.xlsx> "DD/MM/YYYY" "HH:MM"
Updates the consts below + the 'מידע עודכן לפי תאריך' header. Does NOT bump version."""
import json,re,sys,asyncio
from playwright.async_api import async_playwright
P,HR,COMP,EV,DATE,TIME=sys.argv[1:7]
# B448: Datwise's HR export sometimes ignores the "סטטוס: פעיל/ה" filter and returns every status
# (פעיל/ה + לא פעיל/ה + אורח/ת). Keep only active rows by editing the sheet XML in place
# (no openpyxl rewrite, so the native Hebrew headers stay intact for the page parser).
def hr_active_only(path):
    import zipfile,tempfile,os
    z=zipfile.ZipFile(path); names=z.namelist()
    ss=[re.sub('<[^>]+>','',x) for x in re.findall(r'<si>(.*?)</si>',z.read('xl/sharedStrings.xml').decode('utf8'),re.S)] if 'xl/sharedStrings.xml' in names else []
    sh=z.read('xl/worksheets/sheet1.xml').decode('utf8')
    rows=re.findall(r'<row[^>]*>.*?</row>',sh,re.S)
    if not rows: return path
    def cells(r): return {re.match(r'[A-Z]+',m.group(1)).group(0):m.group(0) for m in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*?(?:/>|>.*?</c>)',r,re.S)}
    def val(c):
        v=re.search(r'<v>(.*?)</v>',c,re.S)
        if v: return ss[int(v.group(1))] if 't="s"' in c and ss else v.group(1)
        t=re.search(r'<t[^>]*>(.*?)</t>',c,re.S); return t.group(1) if t else ''
    hdr={k:val(c).strip() for k,c in cells(rows[0]).items()}
    col=next((k for k,v in hdr.items() if v=='סטטוס'),None)
    if not col: return path
    keep=[rows[0]]; drop=0
    for r in rows[1:]:
        c=cells(r).get(col); st=val(c).strip() if c else ''
        if st in ('','פעיל/ה','פעיל'): keep.append(r)
        else: drop+=1
    if not drop: return path
    i0=sh.index(rows[0]); i1=sh.index(rows[-1])+len(rows[-1])
    sh2=sh[:i0]+''.join(keep)+sh[i1:]
    sh2=re.sub(r'<dimension ref="[^"]*"/>','',sh2)
    fd,out=tempfile.mkstemp(suffix='.xlsx'); os.close(fd)
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as w:
        for n in names: w.writestr(n, sh2.encode('utf8') if n=='xl/worksheets/sheet1.xml' else z.read(n))
    print('hr: kept',len(keep)-1,'active rows, dropped',drop,'non-active')
    return out
HR=hr_active_only(HR)
NAMES=['CD_DONE','CD_DONE_LIST','UPCOMING','AT','ST','EMP_START_DATES','CD','MD','BT','FIRSTAID','EVENTS','EVENTS_BY_TYPE','EVENTS_MONTHLY','EVENTS_DIV','EVENTS_STATS']
async def main():
    async with async_playwright() as p:
        b=await p.chromium.launch(); pg=await b.new_page(); errs=[]; pg.on('pageerror',lambda e:errs.append(str(e)))
        await pg.route('**/*',lambda r:r.continue_() if r.request.url.startswith(('file:','data:','blob:')) else r.abort())
        await pg.goto('file://'+P); await pg.wait_for_timeout(3000)
        for k,f in [('attendance',HR),('compliance',COMP),('events',EV)]:
            await pg.set_input_files(f'input[onchange*="\'{k}\'"]',f); await pg.wait_for_timeout(5000)
        vals={n:await pg.evaluate(f"typeof {n}==='undefined'?null:JSON.stringify({n})") for n in NAMES}
        await b.close(); return vals,errs
vals,errs=asyncio.run(main())
s=open(P,encoding='utf-8').read()
def find_end(s,i):
    cl={'{':'}','[':']'}[s[i]]; d=0; q=None; j=i
    while j<len(s):
        c=s[j]
        if q:
            if c=='\\': j+=2; continue
            if c==q: q=None
        elif c in '"\'`': q=c
        elif c in '{[': d+=1
        elif c in '}]':
            d-=1
            if d==0: return j+1
        j+=1
for n,v in vals.items():
    m=re.search(r'const\s+'+n+r'\s*=\s*',s)
    if v is None or not m: print('skip',n); continue
    i=m.end(); s=s[:i]+v+s[find_end(s,i):]
s,c=re.subn(r'מידע עודכן לפי תאריך: \d\d/\d\d/\d{4}(, שעה \d\d:\d\d)?',f'מידע עודכן לפי תאריך: {DATE}, שעה {TIME}',s,count=1)
open(P,'w',encoding='utf-8').write(s)
ST=json.loads(vals['ST'])
print(json.dumps({'hdr':c,'errs':errs,'employees':ST.get('total_employees'),'present':ST.get('present_today'),'completion':ST.get('overall_completion'),'missing':ST.get('total_missing'),'events':len(json.loads(vals['EVENTS']))},ensure_ascii=False))
# B437: GENDER_MAP {eid -> 'זכר'/'נקבה'} from the HR export's "מין" column (for שאילתות search by gender)
try:
    import openpyxl
    ws=openpyxl.load_workbook(HR,read_only=True).active; it=ws.iter_rows(values_only=True); H=[str(x or '').strip() for x in next(it)]
    ie,ig=H.index('מספר עובד/ת'),H.index('מין'); G={}
    for r in it:
        e=str(r[ie] or '').strip().lstrip('0'); g=str(r[ig] or '').strip()
        if e and g in ('זכר','נקבה'): G[e]=g
    s=open(P,encoding='utf-8').read(); new='const GENDER_MAP = '+json.dumps(G,ensure_ascii=False,separators=(',',':'))+';'
    s,n=re.subn(r'const GENDER_MAP = \{.*?\};',lambda m:new,s,count=1,flags=re.S)
    if not n: s=s.replace('const KUPA_MAP = ',new+'\nconst KUPA_MAP = ',1)
    open(P,'w',encoding='utf-8').write(s); print('gender',len(G))
except Exception as ex: print('gender skip',ex)
# B437/B438: EMP_ATTR {eid -> [age band, sector, professional manager, [secondary roles], contractor, company]} from the HR export (no birth dates are stored — only the age band at bake time)
try:
    import openpyxl, datetime as _dt
    ws=openpyxl.load_workbook(HR,read_only=True).active; it=ws.iter_rows(values_only=True); H=[str(x or '').strip() for x in next(it)]
    ie,ib,isc=H.index('מספר עובד/ת'),H.index('תאריך לידה'),H.index('סקטור')
    ipm=H.index('מנהל מקצועי'); ir=[H.index(c) for c in ('תפקיד משני 1','תפקיד משני 2','תפקיד משני 3') if c in H]; ic=H.index('חברת שירות / קבלן'); ico=H.index('חברה')
    today=_dt.date.today(); A={}
    def band(a):
        return 'עד 24' if a<25 else '25–34' if a<35 else '35–44' if a<45 else '45–54' if a<55 else '55–64' if a<65 else '65+'
    for r in it:
        e=str(r[ie] or '').strip().lstrip('0')
        if not e: continue
        b=''; d=str(r[ib] or '').strip()
        try:
            dd=_dt.datetime.strptime(d,'%d/%m/%Y').date(); age=today.year-dd.year-((today.month,today.day)<(dd.month,dd.day))
            if 14<=age<=95: b=band(age)
        except Exception: pass
        sc=str(r[isc] or '').strip()
        g=lambda i:str(r[i] or '').strip()
        r2=[];[r2.append(g(i)) for i in ir if g(i) and g(i) not in r2]
        A[e]=[b,sc,g(ipm),r2,g(ic),g(ico)]
    s=open(P,encoding='utf-8').read(); new='const EMP_ATTR = '+json.dumps(A,ensure_ascii=False,separators=(',',':'))+';'
    s,n=re.subn(r'const EMP_ATTR = \{.*?\};',lambda m:new,s,count=1,flags=re.S)
    if not n: s=s.replace('const KUPA_MAP = ',new+'\nconst KUPA_MAP = ',1)
    open(P,'w',encoding='utf-8').write(s); print('attr',len(A))
except Exception as ex: print('attr skip',ex)
