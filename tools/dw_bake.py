"""Bake native Datwise exports into index.html (via the page's own upload parsers).
usage: python3 tools/dw_bake.py <index.html> <hr.xlsx> <comp.xlsx> <events.xlsx> "DD/MM/YYYY" "HH:MM"
Updates the consts below + the 'מידע עודכן לפי תאריך' header. Does NOT bump version."""
import json,re,sys,asyncio
from playwright.async_api import async_playwright
P,HR,COMP,EV,DATE,TIME=sys.argv[1:7]
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
