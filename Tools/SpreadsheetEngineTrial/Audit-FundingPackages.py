"""Read-only preservation gate for the isolated native Funding trial.

No VBA extraction, passwords or workbook rewrites. Full formula/constant/style
fingerprints come from independent DevExpress manifests, not engine caches.
"""
import argparse
import hashlib
import json
import posixpath
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

ROOT = (Path(__file__).resolve().parents[2] / 'obj/AsposeTrial').resolve()
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

def private(value):
    path = Path(value).resolve(strict=False)
    if ROOT not in path.parents:
        raise ValueError('Private trial paths only')
    return path

def sha(data):
    return hashlib.sha256(data).hexdigest()

def package(path):
    with ZipFile(path) as z:
        names = set(z.namelist())
        rels = {r.attrib['Id']: posixpath.normpath(posixpath.join('xl', r.attrib['Target']))
                for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        sheets = {}
        for s in ET.fromstring(z.read('xl/workbook.xml')).find('s:sheets', NS):
            target = rels[s.attrib[RID]].lstrip('/')
            root = ET.fromstring(z.read(target))
            cells = root.findall('s:sheetData/s:row/s:c', NS)
            counts = {'cellRecords': len(cells), 'metadataCells': sum('cm' in c.attrib for c in cells)}
            for feature in ['dataValidation', 'conditionalFormatting', 'cfRule', 'mergeCell',
                            'drawing', 'legacyDrawing', 'control', 'oleObject', 'tablePart']:
                counts[feature] = len(root.findall('.//s:' + feature, NS))
            counts['arrayAnchors'] = len(root.findall('.//s:f[@t="array"]', NS))
            counts['arrayRanges'] = sorted(f.get('ref', '') for f in root.findall('.//s:f[@t="array"]', NS))
            sheets[s.attrib['name']] = counts
        selected = {n: sha(z.read(n)) for n in sorted(names)
                    if n.startswith('customXml/') or n.endswith('vbaProject.bin') or 'metadata' in n}
        families = {p: sum(n.startswith(p) and not n.endswith('/') for n in names)
                    for p in ['xl/drawings/', 'xl/ctrlProps/', 'xl/activeX/', 'xl/embeddings/', 'xl/charts/', 'customXml/']}
        return {'parts': sorted(names), 'selectedHashes': selected, 'families': families, 'sheets': sheets}

def compare_manifests(a, b):
    fields = ['visibility', 'protected', 'populatedCells', 'formulaCells', 'formulas', 'constants',
              'arrayFormulaRanges', 'arrayFormulas', 'dynamicArrayFormulaRanges', 'dynamicArrayFormulas',
              'populatedCellAppearanceAndLocks']
    targets = {s['name']: s for s in b['sheets']}
    differences = []
    for s in a['sheets']:
        t = targets.get(s['name'], {})
        for f in fields:
            if s.get(f) != t.get(f):
                differences.append({'sheet': s['name'], 'field': f, 'before': s.get(f), 'after': t.get(f)})
    names = [n for n in sorted(a['names'].keys() | b['names'].keys()) if a['names'].get(n) != b['names'].get(n)]
    return {'sheetOrderMatches': [s['name'] for s in a['sheets']] == [s['name'] for s in b['sheets']],
            'nameDifferences': names, 'sheetDifferences': differences,
            'formulaCellsBefore': sum(s['formulaCells'] for s in a['sheets']),
            'formulaCellsAfter': sum(s['formulaCells'] for s in b['sheets'])}

def main():
    p = argparse.ArgumentParser()
    p.add_argument('before'); p.add_argument('after'); p.add_argument('report')
    p.add_argument('--manifests', nargs=2)
    a = p.parse_args()
    source, output, report = map(private, [a.before, a.after, a.report])
    x, y = package(source), package(output)
    features = []
    for sheet, counts in x['sheets'].items():
        for field, value in counts.items():
            other = y['sheets'].get(sheet, {}).get(field)
            if value != other:
                features.append({'sheet': sheet, 'field': field, 'before': value, 'after': other})
    result = {'before': str(source), 'after': str(output),
              'missingParts': sorted(set(x['parts']) - set(y['parts'])),
              'extraParts': sorted(set(y['parts']) - set(x['parts'])),
              'selectedBefore': x['selectedHashes'], 'selectedAfter': y['selectedHashes'],
              'familiesBefore': x['families'], 'familiesAfter': y['families'],
              'worksheetFeatureDifferences': features,
              'limits': 'Feature counts are not semantic equivalence proof. Manifests exclude blank-cell styling and cached result differences.'}
    if a.manifests:
        result['manifestComparison'] = compare_manifests(*(json.loads(private(s).read_text(encoding='utf-8-sig')) for s in a.manifests))
    with report.open('x', encoding='utf-8') as f:
        json.dump(result, f, indent=2)
    print(json.dumps({'report': str(report), 'featureDifferences': len(features),
                      'manifestDifferenceCount': len(result.get('manifestComparison', {}).get('sheetDifferences', [])),
                      'nameDifferences': len(result.get('manifestComparison', {}).get('nameDifferences', [])),
                      'familiesBefore': x['families'], 'familiesAfter': y['families']}, indent=2))

if __name__ == '__main__':
    main()
