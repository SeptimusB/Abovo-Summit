"""Read-only semantic custom-XML/property checks; never emit workbook payloads.

Excel may renumber package parts and change XML whitespace/prefixes. Compare
payload multisets using expanded XML names, not raw ZIP entry hashes. This is
not validation of every package relationship, formatting feature, or VBA run.
"""
import hashlib
import json
import re
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile


def semantic(node):
    return (node.tag, sorted(node.attrib.items()), node.text or "",
            [(semantic(child), child.tail or "") for child in node])


def digest(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=True,
                                    separators=(",", ":")).encode()).hexdigest()


def inspect(path):
    with ZipFile(path) as archive:
        payloads = sorted(digest(semantic(ET.fromstring(archive.read(name))))
                          for name in archive.namelist()
                          if re.fullmatch(r"customXml/item\d+\.xml", name))
        properties = []
        if "docProps/custom.xml" in archive.namelist():
            for prop in ET.fromstring(archive.read("docProps/custom.xml")):
                properties.append((prop.attrib.get("name"),
                                   [semantic(value) for value in prop]))
        return {"customXmlPayloadHashes": payloads,
                "customPropertyHash": digest(sorted(properties))}


baseline = inspect(sys.argv[1])
passed = True
for path in sys.argv[1:]:
    actual = inspect(path)
    matched = actual == baseline
    passed = passed and matched
    print(json.dumps({"file": path, "matchesSource": matched, **actual}))
raise SystemExit(0 if passed else 1)
