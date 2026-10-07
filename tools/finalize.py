"""Bump header version by 1, add changelog row, update Blossom date line.
usage: python3 tools/finalize.py <index.html> "DD/MM/YYYY" "changelog text" ["blossom DD/MM/YYYY, שעה HH:MM"]"""
import re,sys
P,DATE,TXT=sys.argv[1:4]; BL=sys.argv[4] if len(sys.argv)>4 else None
s=open(P,encoding='utf-8').read()
m=re.search(r'(<span style="font-size:12px;font-weight:400;color:#999">\()(\d+)(\)</span>)',s); v=int(m.group(2))+1
s=s[:m.start()]+m.group(1)+str(v)+m.group(3)+s[m.end():]
row=f'<tr style="border-bottom:1px solid #eee"><td style="padding:8px 12px;font-weight:700">{v}</td><td style="padding:8px 12px">{DATE}</td><td style="padding:8px 12px">{TXT}</td></tr>\n            '
k=s.index('<tbody>',s.index('changelogContent'))+len('<tbody>\n            '); s=s[:k]+row+s[k:]
if BL: s=re.sub(r'דוחות אביגל \(בלוסום\) עודכנו: [^·<]*',f'דוחות אביגל (בלוסום) עודכנו: {BL} ',s,count=1)
open(P,'w',encoding='utf-8').write(s); print(v)
