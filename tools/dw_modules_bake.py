"""Bake Datwise 'לוח מחוונים > מודולים' (Reports/Default2.aspx, level=חטיבה) JSON into const DW_MODULES.
Usage: dw_modules_bake.py <index.html> <modules.json>
JSON: {stamp, range, level, cols[15], total[15], rows:[[name, ...15 nums]]}"""
import sys,json,re
p,j=sys.argv[1],sys.argv[2]
d=json.load(open(j,encoding='utf-8'))
s=open(p,encoding='utf-8').read()
new='const DW_MODULES = '+json.dumps(d,ensure_ascii=False,separators=(',',':'))+';'
s,n=re.subn(r'const DW_MODULES = \{.*?\};(?=\n)',lambda m:new,s,count=1,flags=re.S)
assert n==1,'DW_MODULES not found'
open(p,'w',encoding='utf-8').write(s)
print('rows',len(d['rows']),'stamp',d['stamp'])
