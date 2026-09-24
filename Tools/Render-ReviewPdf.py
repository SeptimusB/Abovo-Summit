"""Rasterise the QA PDF using the bundled runtime after Word fallback export."""
import sys
from pathlib import Path
import pypdfium2 as pdfium

pdf = pdfium.PdfDocument(sys.argv[1])
out = Path(sys.argv[2])
out.mkdir(parents=True, exist_ok=True)
for i, page in enumerate(pdf):
    page.render(scale=1.5).to_pil().save(out / f'page-{i+1}.png')
print(f'Rendered {len(pdf)} pages')
