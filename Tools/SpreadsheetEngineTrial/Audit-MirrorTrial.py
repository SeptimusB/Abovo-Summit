"""Read-only mirror-trial evidence. Never export VBA source or rewrite workbooks."""
import argparse
import hashlib
import importlib.util
import json
import posixpath
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

ROOT = (Path(__file__).resolve().parents[2] / 'obj/AsposeTrial').resolve()
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def private(value):
    path = Path(value).resolve()
    if ROOT not in path.parents:
        raise ValueError('Private trial files only')
    return path

def compare_probes(expected, actual):
    result = []
    for name, a in expected.items():
        b = actual[name]
        stats = {'sheet': name, 'nonemptyCompared': 0, 'numericDifferences': 0,
                 'otherDifferences': 0, 'maxNumericDelta': 0}
        for r in range(max(a['rows'], b['rows'])):
            for c in range(max(a['columns'], b['columns'])):
                def cell(block):
                    v = block['values'][r][c] if r < block['rows'] and c < block['columns'] else None
                    return None if v == '' else v
                x, y = cell(a), cell(b)
                if x is None and y is None:
                    continue
                stats['nonemptyCompared'] += 1
                if type(x) in (float, int) and type(y) in (float, int):
                    delta = abs(x-y)
                    stats['maxNumericDelta'] = max(stats['maxNumericDelta'], delta)
                    stats['numericDifferences'] += delta > max(1e-6, abs(x)*1e-10)
                else:
                    stats['otherDifferences'] += x != y
        result.append(stats)
    return result

def vba_review(before, after):
    # Dependency must remain outside the repository. Only hashes and changed
    # attribute NAMES leave memory, never code, passwords or attribute values.
    from oletools.olevba import VBA_Parser
    import difflib
    books = []
    for path in (before, after):
        parser = VBA_Parser(str(path))
        try:
            books.append({(stream, name): code for _, stream, name, code in parser.extract_macros()})
        finally:
            parser.close()
    changes = []
    for key in sorted(set(books[0]) | set(books[1])):
        a, b = books[0].get(key), books[1].get(key)
        if a != b:
            lines = [line[2:] for line in difflib.ndiff((a or '').splitlines(), (b or '').splitlines())
                     if line[:2] in ('- ', '+ ')]
            changes.append({'module': key[1], 'changedLineCount': len(lines),
                            'onlyFormBaseAttribute': a is not None and b is not None and key[1].endswith('.frm')
                            and all(line.startswith('Attribute VB_Base = ') for line in lines)})
    return {'moduleCountBefore': len(books[0]), 'moduleCountAfter': len(books[1]),
            'changes': changes,
            'executableSourceUnchangedExceptNativeFormBaseAttributes': bool(books[0]) and all(c['onlyFormBaseAttribute'] for c in changes),
            'limitation': 'Does not certify form designer storage, compiled code, references or UI operation.'}

def formula_geometry(path):
    with ZipFile(path) as z:
        book = ET.fromstring(z.read('xl/workbook.xml'))
        rels = {r.get('Id'): posixpath.normpath(posixpath.join('xl', r.get('Target'))).lstrip('/')
                for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        sheets = {}
        for sheet in book.find('s:sheets', NS):
            entry = rels[sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')]
            digest = hashlib.sha256()
            count = 0
            with z.open(entry) as stream:
                for _, cell in ET.iterparse(stream, events=('end',)):
                    if cell.tag == '{'+NS['s']+'}c':
                        if cell.find('s:f', NS) is not None:
                            digest.update((cell.get('r')+'\n').encode())
                            count += 1
                        cell.clear()
            sheets[sheet.get('name')] = (count, digest.hexdigest())
        names = {(n.get('name'), n.get('localSheetId')): n.text
                 for n in book.findall('s:definedNames/s:definedName', NS)}
        return sheets, names

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('trial_report')
    parser.add_argument('excel_reference')
    parser.add_argument('output')
    args = parser.parse_args()
    report_path, reference_path, output = map(private, (args.trial_report, args.excel_reference, args.output))
    report = json.loads(report_path.read_text(encoding='utf-8-sig'))
    reference = json.loads(reference_path.read_text(encoding='utf-8-sig'))
    if not report['harnessCompleted'] or not reference['success'] or report['inputHash'] != reference['inputHash']:
        raise ValueError('Need completed trials with the identical input hash')
    baseline = report_path.parent / 'baseline.xlsm'
    if not baseline.exists():
        baseline = private(report['input'])  # Native-only trace uses its immutable private input directly.
    candidate = private(report['tests'].get('structuralReadbackFile') or report['structuralReadbackFile'])
    oracle = private(reference['output'])
    spec = importlib.util.spec_from_file_location('funding_package', Path(__file__).with_name('Audit-FundingPackages.py'))
    package = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(package)
    a, b = package.package(oracle), package.package(candidate)
    formula_a, names_a = formula_geometry(oracle)
    formula_b, names_b = formula_geometry(candidate)
    features = []
    for name, counts in a['sheets'].items():
        for key in ('metadataCells', 'arrayAnchors', 'arrayRanges'):
            if counts[key] != b['sheets'].get(name, {}).get(key):
                features.append({'sheet': name, 'feature': key})
    data = {
        'trial': str(report_path), 'oracle': str(oracle),
        'probeComparison': compare_probes(reference['calculatedProbes'], report['tests']['correctedGearProbes']),
        'arrayMetadataDifferencesAgainstExcelVbaOracle': features,
        'formulaAddressDifferences': [name for name in sorted(set(formula_a) | set(formula_b)) if formula_a.get(name) != formula_b.get(name)],
        'oracleFormulaCellCount': sum(value[0] for value in formula_a.values()),
        'candidateFormulaCellCount': sum(value[0] for value in formula_b.values()),
        'definedNameDifferences': [{'name': key[0], 'scope': key[1]} for key in sorted(set(names_a) | set(names_b), key=repr) if names_a.get(key) != names_b.get(key)],
        'oracleFamilies': a['families'], 'candidateFamilies': b['families'],
        'vba': vba_review(baseline, candidate),
        'limits': 'Five calculated sheet ranges and array geometry; full formulas, formats, UI and financial approval are separate gates.'
    }
    with output.open('x', encoding='utf-8') as stream:
        json.dump(data, stream, indent=2)
    print(json.dumps(data, indent=2))

if __name__ == '__main__':
    main()
