"""Read-only small formula samples from private trial packages; never reads VBA."""
import sys, json, posixpath, re
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

root=(Path(__file__).resolve().parents[2]/'obj/AsposeTrial').resolve()
path=Path(sys.argv[1]).resolve()
if root not in path.parents: raise ValueError('Private trial file required')
ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
rid='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
with ZipFile(path) as z:
    rels={r.get('Id'):posixpath.normpath(posixpath.join('xl',r.get('Target'))).lstrip('/') for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    result=[]; udf_count={}
    for sheet in ET.fromstring(z.read('xl/workbook.xml')).find('s:sheets',ns):
        name=sheet.get('name'); doc=ET.fromstring(z.read(rels[sheet.get(rid)]))
        for c in doc.findall('s:sheetData/s:row/s:c',ns):
            f=c.find('s:f',ns)
            if f is None or not f.text: continue
            matches=re.findall(r'([\w\.]*?(?:PMCost|RespCost))\(',f.text,re.I)
            for match in matches:
                udf_count[match]=udf_count.get(match,0)+1
                if udf_count[match]<=2: result.append({'sheet':name,'cell':c.get('r'),'formula':f.text})
            if name in ['Detailed Comp Inc - Trad View','Development Expenditure'] and c.get('r') in ['A1','AV417']:
                result.append({'sheet':name,'cell':c.get('r'),'formula':f.text})
    print(json.dumps({'udfCount':udf_count,'samples':result},indent=2))
