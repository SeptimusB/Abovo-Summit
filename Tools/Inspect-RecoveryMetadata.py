"""Read-only diagnosis of dynamic-array metadata in disposable XLSM outputs."""
import sys, json, zipfile, xml.etree.ElementTree as ET
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
for path in sys.argv[1:]:
    with zipfile.ZipFile(path) as z:
        rows=[]
        for entry in z.namelist():
            if not entry.startswith('xl/worksheets/') or not entry.endswith('.xml'): continue
            with z.open(entry) as content:
                for _, cell in ET.iterparse(content, events=('end',)):
                    if cell.tag != NS+'c': continue
                    f=cell.find(NS+'f')
                    if 'cm' in cell.attrib or (f is not None and f.get('t')=='array'):
                        rows.append({'part':entry,'cell':cell.get('r'),'cm':cell.get('cm'),'formula':None if f is None else f.text,'attributes':{} if f is None else f.attrib})
                    cell.clear()
        print(json.dumps({'file':path,'metadata':z.read('xl/metadata.xml').decode() if 'xl/metadata.xml' in z.namelist() else None,'cells':rows}))
