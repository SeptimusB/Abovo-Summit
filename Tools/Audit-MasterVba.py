"""Read every VBA module; persist inventory/metadata only, never raw source.

oletools must be installed outside this repository. Macros are never executed.
Optional --view lists sanitized source to stdout for review, not into artifacts.
"""
import argparse
import hashlib
import json
import re
import fnmatch
from pathlib import Path
from oletools.olevba import VBA_Parser

START = re.compile(r'^\s*(?:(?:Public|Private|Friend|Static)\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)', re.I)
END = re.compile(r'^\s*End\s+(?:Sub|Function|Property)\s*$', re.I)
STRING = re.compile(r'"(?:[^"]|"")*"')


def code_only(line):
    quoted = False
    i = 0
    while i < len(line):
        if line[i] == '"':
            if quoted and i+1 < len(line) and line[i+1] == '"':
                i += 2
                continue
            quoted = not quoted
        elif line[i] == "'" and not quoted:
            return line[:i]
        i += 1
    return line


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('workbook')
    ap.add_argument('--output')
    ap.add_argument('--view', nargs='*')
    args = ap.parse_args()
    parser = VBA_Parser(args.workbook)
    modules = []
    try:
        for _, stream, name, source in parser.extract_macros():
            lines = source.splitlines()
            procedures = []
            sensitive_lines = set()
            # Credentials may be compared on a later line than the password
            # prompt. Redact literals for the WHOLE sensitive procedure.
            start = 0
            for idx, line in enumerate(lines):
                if START.match(code_only(line)):
                    start = idx
                if END.match(code_only(line)):
                    sensitive = [code_only(s) for s in lines[start:idx+1]
                        if not re.search(r'Password\s*:=\s*PW\w*\b', code_only(s), re.I)]
                    if any(STRING.search(s) and re.search(r'password|\bpwd\b|\bpw\b|RejData', s, re.I) for s in sensitive):
                        sensitive_lines.update(range(start+1, idx+2))
            current = None
            for number, line in enumerate(lines, 1):
                code = code_only(line)
                match = START.match(code)
                if match:
                    current = {'name': match[2], 'kind': match[1], 'start': number, 'lines': []}
                if current is not None:
                    current['lines'].append(code)
                    if END.match(code):
                        body = '\n'.join(current.pop('lines'))
                        current['end'] = number
                        current['sha256'] = hashlib.sha256(body.encode()).hexdigest()
                        tokens = STRING.sub('""', body)
                        current['_tokens'] = '\n'.join(tokens.splitlines()[1:])
                        current['flags'] = [label for label, pattern in {
                            'insert': r'\.Insert\b', 'delete': r'\.Delete\b',
                            'copy_paste': r'\.(?:Copy|Paste|PasteSpecial|FillDown|FillRight)\b',
                            'calculate': r'\b(?:Calculate|CalculateFullRebuild|ABVCalculate)\b',
                            'sheet_group': r'\bSheets\s*\(', 'state': r'\b(?:EnableEvents|Calculation|ScreenUpdating)\b',
                            'external_io': r'\b(?:Open|SaveAs|SaveCopyAs|FileCopy|Kill|Shell)\b',
                            'event_entry': r'\b(?:Worksheet_|Workbook_)',
                            'sync': r'\b\w*(?:Synchronise|TransDB|TransactionList)\w*\b',
                            'clear_inputs': r'\bClearContents\b',
                            'resize_names': r'\b(?:Resize|Names\.Add)\b',
                            'broad_error_suppression': r'\bOn\s+Error\s+Resume\s+Next\b',
                            'vba_project_access': r'\b(?:VBProject|VBComponents|CodeModule)\b',
                        }.items() if re.search(pattern, tokens, re.I)]
                        current['named_ranges'] = sorted(set(re.findall(r'\bRange\s*\(\s*"([A-Za-z_]\w*)"', body, re.I)))
                        current['worksheets'] = sorted(set(re.findall(r'\b(?:Sheets|Worksheets)\s*\(\s*"([^"\r\n]+)"', body, re.I)))
                        current['calls'] = sorted(set(re.findall(r'\bCall\s+(\w+)', tokens, re.I)))
                        procedures.append(current)
                        current = None
            modules.append({'name': name, 'stream': stream, 'line_count': len(lines),
                            'sha256': hashlib.sha256(source.encode()).hexdigest(), 'procedures': procedures,
                            'incomplete_procedure': None if current is None else current['name']})
            if args.view is not None and any(fnmatch.fnmatchcase(name, pattern) for pattern in args.view):
                print('\nMODULE', name, 'FULL EXECUTABLE TEXT (comments/attributes omitted; credentials redacted)')
                for number, line in enumerate(lines, 1):
                    code = code_only(line).rstrip()
                    if not code.strip() or code.lstrip().startswith(('Attribute ', 'VERSION ', 'BEGIN', 'END')):
                        continue
                    if number in sensitive_lines or name.startswith(('Protection_', 'ProtectForm', 'UnprotectForm')) or re.search(r'password|\bpwd\b|\bpw\b|RejData', code, re.I):
                        code = STRING.sub('"<redacted>"', code)
                    print(f'{number}: {code}')
    finally:
        parser.close()
    symbols = {}
    for module in modules:
        for proc in module['procedures']:
            symbols.setdefault(proc['name'].lower(), []).append(module['name']+':'+proc['name'])
    for module in modules:
        for proc in module['procedures']:
            tokens = proc.pop('_tokens')
            # A conservative lexical candidate graph, not a VBA compiler.
            proc['project_call_candidates'] = sorted({target for token in re.findall(r'\b\w+\b', tokens)
                for target in symbols.get(token.lower(), [])
                if target != module['name']+':'+proc['name']})
    report = {'workbook': str(Path(args.workbook).resolve()),
              'sha256': hashlib.sha256(Path(args.workbook).read_bytes()).hexdigest(),
              'module_count': len(modules), 'line_count': sum(m['line_count'] for m in modules),
              'procedure_count': sum(len(m['procedures']) for m in modules),
              'scope': 'Complete source inventory and static scan, not VBA execution or financial sign-off.',
              'contains_raw_vba': False, 'modules': modules}
    if args.output:
        Path(args.output).write_text(json.dumps(report, indent=2)+'\n', encoding='utf-8')
    if args.view is None:
        print(json.dumps({k: v for k, v in report.items() if k != 'modules'}, indent=2))


if __name__ == '__main__':
    main()
