"""Read-only VBA source comparison; emit hashes/counts, never source or passwords.

Requires oletools installed OUTSIDE the repository (e.g. a temporary PYTHONPATH).
Pass the original workbook, Summit-saved test result and Excel-saved test result.
Does not execute VBA or extract source to disk.
"""
import hashlib
import sys
from oletools.olevba import VBA_Parser


def hashes(path):
    parser = VBA_Parser(path)
    try:
        result = {}
        for _, stream, name, code in parser.extract_macros():
            key = (stream, name)
            if key in result:
                raise ValueError("Duplicate module identity")
            payload = code.encode("utf-8") if isinstance(code, str) else code
            result[key] = hashlib.sha256(payload).hexdigest()
        if not result:
            raise ValueError("No VBA modules found")
        return result
    finally:
        parser.close()


if __name__ == "__main__":
    baseline = None
    for path in sys.argv[1:]:
        actual = hashes(path)
        print(path, "modules=", len(actual), "aggregate=",
              hashlib.sha256(repr(sorted(actual.items())).encode()).hexdigest())
        if baseline is None:
            baseline = actual
        elif actual != baseline:
            raise SystemExit("FAIL: VBA module source differs")
    if len(sys.argv) < 3:
        raise SystemExit("Supply at least two XLSB paths")
    print("PASS: all module identities and source hashes identical; VBA execution not tested.")
