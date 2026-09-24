"""Revise Jon's annotated v01, retaining screenshots and untouched source bytes."""
import argparse
import re
from copy import deepcopy
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

ORANGE = 'C45D18'
STATUS = '> Alex to test'

def replace(p, text, orange=False):
    drawings = [deepcopy(r._r) for r in p.runs if r._r.xpath('.//w:drawing')]
    for node in drawings:
        for t in node.xpath('.//w:t'):
            t.getparent().remove(t)
    p.clear()
    r = p.add_run(text)
    r.font.color.rgb = RGBColor.from_string(ORANGE if orange else '333333')
    for node in drawings:
        p.add_run().add_break()
        p._p.append(node)

def after(p, text, orange=False):
    el = OxmlElement('w:p')
    p._p.addnext(el)
    q = Paragraph(el, p._parent)
    q.style = 'Normal'
    q.paragraph_format.left_indent = Pt(14)
    q.paragraph_format.space_after = Pt(7)
    replace(q, text, orange)
    return q

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('source', type=Path)
    ap.add_argument('output', type=Path)
    a = ap.parse_args()
    if a.output.exists():
        raise SystemExit('Output already exists; choose the next revision')
    original = a.source.read_bytes()
    d = Document(a.source)
    ps = {i + 1:p for i,p in enumerate(d.paragraphs)}
    assert len(ps) == 206 and 'version 01' in ps[1].text
    assert 'Management' in ps[70].text and 'Add lines' in ps[205].text
    replace(ps[1], 'Summit client issue review version 03')
    ps[1].style = 'Title'
    replace(ps[2], 'Ready to test with Summit 2.84\n23 September 2026\n\nAlex, please retest the orange items and answer the questions beneath the remaining issues. This is the working review for the next test round. The original issue order and your screenshots are retained.')
    replace(ps[3], 'Yellow highlights fixes awaiting testing. Orange issue text means believed fixed, not accepted. Green retains the previous agreed status and does not prove a fix. Unhighlighted questions and partial fixes remain open. Testing in Summit does not replace the accountant\'s financial checks.')
    replace(ps[4], 'Client issues from Summit 2 51')
    # Normalise status wording, without replacing drawings or the original issue text.
    for i,p in ps.items():
        if i <= 4 or p._p.xpath('.//w:drawing'):
            continue
        t = p.text
        if 'Alex to confirm' in t:
            t = re.sub(r'>?\s*[\u201c\u201d\"]?Alex to confirm(?: on his PC)?[\u201c\u201d\"]?\.?', STATUS, t)
        t = t.replace('Please confirm in the client retest.', STATUS).replace('Please confirm in retest.', STATUS)
        if t != p.text:
            replace(p, t, orange=True)
    updates = {
        20: 'The name comes from SelectTrust in the workbook. Changing Setup Company Name refreshes the open headings, main file tab and FileInstance summary after the edit is committed. Undo and Redo refresh them too.\n' + STATUS,
        27: 'Repeating year headings now display Year 1, Year 2, etc. Stored years remain numeric; financial-year suffixes remain available.\n' + STATUS,
        29: 'Return uses a left arrow at the far right with its own separator. History uses the former curved Return icon. The Return button and separator appear only when there is an interface to return to.\n' + STATUS,
        35: 'Check Sheet is under Global Assumptions. An imbalance is a soft notice: normal Save and Save As remain available. Recovery backup defaults to continuing while figures are being entered; an existing choice to pause is retained.\n' + STATUS,
        38: 'Believed fixed: the company-name editor now spans more of the available width. The requested change was width, not a change to input validation.\n' + STATUS,
        40: 'A remembered Check Sheet result is checked automatically on open without a popup or an initial warning heading. Unsaved test findings stay in the current session. A successful Save records the saved state; a recovery save records the recovery copy\'s state. Inspection alone does not write XML to the source workbook.\n' + STATUS,
        41: 'New trial: Options > Check Sheet trial can calculate just that worksheet after the selected interval and idle period. It updates the balance message only when the state changes. Temporary timings cover the worksheet calculation and individual typed edits. This is separate from the full Integrity checks for formula errors, broken references and structural problems.',
        42: 'Please test the unsaved-break/discard/reopen sequence, saved-error/reopen sequence, and Yes/No override clearing with Undo/Redo. Turn off temporary timings when the trial is accepted. The reminder has been arranged.\n' + STATUS,
        45: 'Believed fixed: Stock Description has a larger minimum width, equivalent to 42 characters, and remains resizable. Please confirm its width at the client\'s normal display scale.\n' + STATUS,
        47: 'Believed fixed: an IsDummy read-only separator column has been added in the interface XML between the requested stock groups. It does not insert a column into the workbook.\n' + STATUS,
        48: '', 49: '',
        56: 'Formatting on Click to access Economic Assumptions is indented.\nBelieved fixed: leading XML whitespace is trimmed and the text uses the available section width.\n' + STATUS,
        57: 'Summary Categories tab link button does not show all its text.\nBelieved fixed: the button can size to its caption and use the section width. Check on the client\'s PC.\n' + STATUS,
        62: 'Believed fixed: the same shared whitespace and width changes apply here.\n' + STATUS,
        63: 'Link to inflationary increases does not show all its text.\nBelieved fixed: the shared link-button sizing change applies here.\n' + STATUS,
        65: 'Believed fixed: Add Lines is disabled for Summary Other Income Category. The table has no supported expansion action in the master.\n' + STATUS,
        66: 'Agreed scope: remove expansion from this interface definition; do not add a new workbook expansion rule.',
        71: 'Summary Categories tab read-only cells are difficult to read, with white text on a pale background.',
        72: 'Believed fixed: the read-only drawing path now uses the source cell\'s colours instead of forcing white text on lavender. The screenshots below show the reported problem.\n' + STATUS,
        76: 'Partly repaired: the two category fields now refresh their cached source colours and rule state, rather than retaining stale colours. Full conditional-format parity with Excel is still open. Question: after clearing Description and committing the edit, which colour or editability still differs from the corresponding Excel cell?',
        78: 'Confirmed decision: keep the Year 1 heading read-only and the amounts beneath it editable. No extra ROInitialLines value has been added: that would lock the next year selector, not the amount cells.\n' + STATUS,
        82: 'Believed fixed: Description and both category columns are pinned on the left in the Management Costs tables. They remain subject to the existing editability rules.\n' + STATUS,
        83: 'Agreed scope: pin the category columns as well as Description.',
    }
    non_orange = {3,4,41,66,76,83}
    for i,t in updates.items():
        replace(ps[i], t, orange=i not in non_orange)
    # The original paste grouped one status with the following issue.
    replace(ps[34], 'Validation Check' + ps[34].text.split('Validation Check', 1)[1], orange=True)
    for p in d.paragraphs:
        if STATUS in p.text and not p._p.xpath('.//w:drawing'):
            body = p.text.replace(STATUS, '').strip()
            replace(p, (body + '\n' if body else '') + STATUS, orange=True)
    for p in list(d.paragraphs):
        if not p.text.strip() and not p._p.xpath('.//w:drawing'):
            p._p.getparent().remove(p._p)
    for shape in d.inline_shapes:
        if shape.height > Inches(3.2):
            shape.width = int(shape.width * Inches(3.2) / shape.height)
            shape.height = Inches(3.2)
    for i in [6,36,43,51,55,60,67,70,84,92,97,122,130,142,145,163,170,176,180,186,197,201]:
        ps[i].paragraph_format.keep_with_next = True
    ps[72].paragraph_format.keep_with_next = True
    after(ps[21], 'Question: which current button is blank? Please identify its neighbours or attach a current screenshot.')
    # Keep genuine outstanding work distinct from retest requests.
    for i in [7,9,11,13,15,17,19,22,24,26,28,30,32,37,44,46,56,57,61,63,64,71,77,81]:
        for r in ps[i].runs:
            r.font.color.rgb = RGBColor.from_string(ORANGE)
    after(ps[206], 'Ready to test checks for this round', False).style = 'Heading 1'
    d.add_paragraph('Private-copy native tests cover Yes/No entry and Undo/Redo, session-only Check Sheet findings, successful-save persistence, recovery policy, automatic recheck and idle scheduling, pinned Management Costs descriptors, source-colour refresh and the wider company-name editor. The original master files are unchanged.')
    d.add_paragraph('Outstanding priorities are conditional-format refresh across the other interfaces, nested scrolling and layouts, the ambiguous date and dropdown requests, and the questions retained beneath each issue. These are not marked fixed by the balance-check changes.')
    d.add_paragraph('Suggested test: make an imbalance without saving, check it, close and discard, then reopen. Repeat with a deliberate Save. Confirm the balance message follows the file that was actually saved. Test normal data-entry imbalances separately from formula errors; the latter must remain Integrity findings.')
    for name in ['Title','Subtitle','Heading 1','Heading 2','Heading 3']:
        d.styles[name].font.color.rgb = RGBColor(0,0,0)
    for p in d.paragraphs:
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_together = True
        if STATUS in p.text:
            for r in p.runs:
                if not r.text:
                    continue
                r.font.color.rgb = RGBColor.from_string('222222')
                shade = OxmlElement('w:shd')
                shade.set(qn('w:val'), 'clear')
                shade.set(qn('w:fill'), 'FFFF99')
                r._r.get_or_add_rPr().append(shade)
        if p.style.name.startswith(('Title','Heading','Subtitle')):
            p.paragraph_format.keep_with_next = True
            for r in p.runs:
                r.font.color.rgb = RGBColor(0,0,0)
                r.font.underline = False
        p.paragraph_format.widow_control = True
    d.save(a.output)
    assert sha256(a.source.read_bytes()).digest() == sha256(original).digest()
    with ZipFile(a.source) as src, ZipFile(a.output) as out:
        media = [n for n in src.namelist() if n.startswith('word/media/')]
        assert all(src.read(n) == out.read(n) for n in media)
    print(f'Created {a.output}; source unchanged; {len(media)} original images preserved')

if __name__ == '__main__':
    main()
