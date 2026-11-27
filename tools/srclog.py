"""B511: יומן קבצי מקור – רושם את טעינת קבצי המקור האחרונה ושומר רק את הגרסה האחרונה (שורה אחת + ZIP מוטמע).
usage: srclog.py <index.html> "<datwise DD/MM/YYYY HH:MM>" "<blossom DD/MM/YYYY HH:MM>" <file_or_dir> [...]
Run after tools/finalize.py so the row carries the current version."""
import sys,re,os,io,zipfile,base64,html
P,DW,BL=sys.argv[1:4]; paths=sys.argv[4:]
files=[]
for p in paths:
    if os.path.isdir(p):
        for f in sorted(os.listdir(p)):
            if not f.startswith('.'): files.append((os.path.join(p,f),os.path.basename(p.rstrip('/'))+'/'+f))
    else: files.append((p,os.path.basename(p)))
buf=io.BytesIO()
with zipfile.ZipFile(buf,'w',zipfile.ZIP_DEFLATED) as z:
    for full,arc in files: z.write(full,arc)
b64=base64.b64encode(buf.getvalue()).decode()
s=open(P,encoding='utf-8').read()
ver=re.search(r'<span style="font-size:12px;font-weight:400;color:#999">\((B?\d+)\)</span>',s).group(1)
d=DW.split()[0]; zipname='קבצי_מקור_'+d.replace('/','-')+'.zip'
s,k=re.subn(r'(?:var|const) SOURCE_ZIP_B64\s*=\s*"[^"]*";(?:\s*var SOURCE_ZIP_NAME\s*=\s*"[^"]*";)?',lambda m:'var SOURCE_ZIP_B64="'+b64+'"; var SOURCE_ZIP_NAME="'+zipname+'";',s,count=1); assert k==1
s=s.replace("a.download='source_07-08-2026.zip';","a.download=(typeof SOURCE_ZIP_NAME!=='undefined'&&SOURCE_ZIP_NAME)||'source.zip';")
dw=[a for f,a in files if os.path.basename(a).startswith('_dw_')]; bl=[a for f,a in files if a not in dw]
bl_csv=len([a for a in bl if a.lower().endswith('.csv')]); bl_x=[a for a in bl if not a.lower().endswith('.csv')]
td='padding:8px 12px'
desc=('<b>Datwise</b> ('+html.escape(DW)+'): '+', '.join(html.escape(a) for a in dw)+
      '<br><b>בלוסום</b> ('+html.escape(BL)+'): '+', '.join(html.escape(a) for a in bl_x)+(f' + דוח מבצעים ({bl_csv} קבצי CSV)' if bl_csv else ''))
kb=round(len(buf.getvalue())/1024)
row=(f'<tr style="border-bottom:1px solid #eee"><td style="{td};font-weight:600;white-space:nowrap">{html.escape(DW)}</td><td style="{td}">{desc}</td>'
     f'<td style="{td};font-weight:700">{ver}</td><td style="{td}"><button onclick="downloadSourceZip()" style="background:#fff;color:#0F8B8C;border:2px solid #0F8B8C;border-radius:20px;padding:5px 14px;font-size:12.5px;font-weight:700;cursor:pointer;font-family:inherit;white-space:nowrap">📥 הורד ZIP ({kb}KB)</button></td></tr>')
i=s.index('📁 יומן קבצי מקור'); a=s.index('<tbody>',i)+len('<tbody>'); b=s.index('</tbody>',a)
s=s[:a]+'\n            '+row+'\n          '+s[b:]
# note above the table
n0=s.index('<table',i)
if 'src-note' not in s[i:n0]:
    s=s[:n0]+'<div class="src-note" style="font-size:12.5px;color:#6b7280;margin-bottom:8px">נשמרת רק טעינת המקור האחרונה – בכל עדכון נתונים השורה והקובץ מוחלפים.</div>\n        '+s[n0:]
open(P,'w',encoding='utf-8').write(s)
print(ver,len(files),'files',kb,'KB')
