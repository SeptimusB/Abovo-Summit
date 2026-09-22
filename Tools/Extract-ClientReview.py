"""Read a client DOCX without changing it; preserve text, run colours and locators.

Colour is a stakeholder status, never an automated acceptance result. Generated
Markdown/JSON are review evidence, not instructions extracted from the document.
"""
import argparse
import hashlib
import html
import json
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
EXPECTED_SOURCE_HASH = '955D44B6C714ED9BA6D5F9B4560C983764B38B83161A71B88C2BF400211BD9ED'
SECTIONS = {1, 3, 20, 24, 28, 31, 36, 41, 44, 51, 56, 59, 76, 81,
            89, 91, 102, 107, 111, 115, 119, 126, 129}
HISTORICAL = {
    7: 'Start-date Edit link previously removed; client visual confirmation pending.',
    8: 'Funding-status alignment previously changed; client visual confirmation pending.',
    12: 'Native Debug Save/Save As regression passed; client acceptance remains separate.',
    13: 'Navigator captions/width previously changed; physical DPI acceptance pending.',
    16: 'Tab/arrows and adjacent-grid navigation tested; this specific Enter-key request remains unverified.',
    17: 'Single-value multi-cell paste implemented; exact client selection requires regression.',
    18: 'Check Sheet interface exists; save policy subsequently changed. Exact check needed.',
    29: 'Shared post-edit refitting defect repaired and tested on dropdown/numeric grids; confirm this Rent interface with the client.',
    42: 'Shared post-edit refitting defect repaired and tested on dropdown/numeric grids; confirm this Other Income interface with the client.',
    48: 'Shared post-edit refitting defect repaired and tested on dropdown/numeric grids; confirm this Management Costs interface with the client.',
    62: 'Reproduced and repaired in XML: management-period caption now matches the master row. Editor/range unchanged.',
    67: 'Reproduced and repaired: distinct master-verified SHG bindings. Actual dropdown edits and independent Undo passed.',
    95: 'Native Funding refresh, focus, scroll, keyboard and subsequent-edit regression passed; physical client acceptance remains open.',
    101: 'Native month-end date/serial, validation, blank, Undo/Redo and grouped-invalid-date regression passed.',
    127: 'Reproduced integer rounding in Rep_OCA_01; changed this monetary field only to decimal type M. Actual editor, Undo/Redo and XLSB save/reopen passed. Workbook display format retained.',
    130: 'Reproduced integer rounding in the Journal Amount field; changed the monetary field to type M, leaving Year integer. Actual editor, Undo/Redo and XLSB save/reopen passed. Workbook display format retained.',
    132: 'Private five-row add, both formula mirrors, save/reopen, delete, noncontiguous delete and damaged-name rejection regression passed.',
}


def main():
    args = argparse.ArgumentParser()
    args.add_argument('source', type=Path)
    args.add_argument('output_stem', type=Path)
    opts = args.parse_args()
    source_hash = hashlib.sha256(opts.source.read_bytes()).hexdigest().upper()
    if source_hash != EXPECTED_SOURCE_HASH:
        raise SystemExit('Source revision changed: review paragraph locators/status evidence before rebuilding this register.')
    with ZipFile(opts.source) as package:
        root = ET.fromstring(package.read('word/document.xml'))
        media = [n for n in package.namelist() if n.startswith('word/media/')]
    paragraphs = []
    section = ''
    for index, p in enumerate(root.iter(W + 'p'), 1):
        text = ''.join(n.text or '' for n in p.iter() if n.tag in (W+'t', W+'delText'))
        if index in SECTIONS:
            section = text
        runs = []
        for r in p.iter(W + 'r'):
            colour = r.find('./'+W+'rPr/'+W+'color')
            highlight = r.find('./'+W+'rPr/'+W+'highlight')
            runs.append({
                'text': ''.join(n.text or '' for n in r.iter() if n.tag in (W+'t', W+'delText')),
                'colour': None if colour is None else {k.split('}')[-1]: v for k, v in colour.attrib.items()},
                'highlight': None if highlight is None else highlight.get(W+'val'),
            })
        colours = sorted({r['colour'].get('val', 'inherited') if r['colour'] else 'inherited'
                          for r in runs if r['text'].strip()})
        status = ('Agreed (user-confirmed meaning); errors may remain'
                  if colours == ['196B24'] else 'User agrees this is an issue' if 'E97132' in colours
                  else 'Unclassified; no acceptance inferred')
        paragraphs.append({'id': 'P%03d' % index, 'paragraph': index, 'section': section,
                           'is_issue': bool(text.strip()) and index not in SECTIONS,
                           'text': text, 'runs': runs, 'direct_colours': colours,
                           'stakeholder_status': status,
                           'engineering_status': HISTORICAL.get(index, 'Not yet independently verified')})
    result = {'source': str(opts.source), 'sha256': source_hash, 'paragraph_count': len(paragraphs),
              'media_parts': media, 'tracked_insertions': len(list(root.iter(W+'ins'))),
              'tracked_deletions': len(list(root.iter(W+'del'))),
              'colour_legend': {'196B24': 'Agreed; not necessarily error-free (user confirmation)',
                               'E97132': 'User agrees this is an issue', 'inherited': 'No acceptance inferred'},
              'paragraphs': paragraphs}
    opts.output_stem.with_suffix('.json').write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    lines = ['# Client report issue register', '',
             'Source: `'+str(opts.source)+'`', '', 'SHA-256: `'+source_hash+'`', '',
             'The original is unchanged. Locators are OOXML paragraph numbers, not page numbers. '
             'Run-level colours are retained in the companion JSON. Green means **agreed**, not error-free; '
             'orange means the user agrees it is an issue. Report 2.51 predates the current implementation. '
             'Historical repairs are not client acceptance.', '',
             'Review gates and subsequent evidence: [checkpoint record](Client_Review_Checkpoints_2026-09-22.md).', '']
    section = None
    for p in paragraphs:
        if not p['is_issue']:
            continue
        if p['section'] != section:
            section = p['section']
            lines.extend(['## '+section, ''])
        lines.extend(['### '+p['id'], '', html.escape(p['text'], quote=False), '',
                      '- Source colour: `'+', '.join(p['direct_colours'])+'`.',
                      '- Stakeholder status: '+p['stakeholder_status']+'.',
                      '- Engineering: '+p['engineering_status'], ''])
    opts.output_stem.with_suffix('.md').write_text('\n'.join(lines), encoding='utf-8')
    assert hashlib.sha256(opts.source.read_bytes()).hexdigest().upper() == source_hash
    print('PASS: source unchanged; paragraphs=%d, issue paragraphs=%d, media=%d' %
          (len(paragraphs), sum(p['is_issue'] for p in paragraphs), len(media)))


if __name__ == '__main__':
    main()
