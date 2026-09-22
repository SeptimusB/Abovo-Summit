"""Read-only recovery/custom XML preservation assertions; no workbook writes."""
import sys, collections, xml.etree.ElementTree as ET
from zipfile import ZipFile

HISTORY = 'urn:abovo:summit:recovery-history:1'
def tree(node):
    text = node.text or ''
    if len(node) and not text.strip(): text = ''  # Serializer indentation, not leaf data.
    return (node.tag, tuple(sorted(node.attrib.items())), text, tuple(tree(c) for c in node))
def inspect(path):
    payloads = []
    properties = {}
    with ZipFile(path) as package:
        for name in package.namelist():
            if name.startswith('customXml/') and name.endswith('.xml'):
                root = ET.fromstring(package.read(name))
                if root.tag.endswith('}datastoreItem'): continue
                payloads.append(tree(root))
        for prop in ET.fromstring(package.read('docProps/custom.xml')):
            properties[prop.get('name')] = tuple(tree(c) for c in prop)
    return collections.Counter(payloads), properties

baseline, original_props = inspect(sys.argv[1])
history = None
for path in sys.argv[2:]:
    actual, props = inspect(path)
    assert not baseline - actual, 'Existing custom XML payload changed: '+path
    found = [value for value in actual.elements() if value[0] == '{'+HISTORY+'}RecoveryHistory']
    assert len(found) == 1, 'Expected one history part'
    if history is None: history = found[0]
    assert history == found[0], 'Recovered audit history changed'
    for key, value in original_props.items():
        if key == 'Abovo.Summit.ResultsPending': continue
        assert props.get(key) == value, 'Original document property changed: '+key
    print('PASS: existing custom XML/properties and one unchanged recovery history:', path)
