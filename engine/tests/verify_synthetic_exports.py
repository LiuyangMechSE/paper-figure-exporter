"""Independent end-to-end checks for default synthetic fixture exports.

Usage: python tests/verify_synthetic_exports.py path/to/output
The fixture should yield four numbered PDFs; uncertain/hidden/off-slide cases
must be reported rather than exported by default.
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from PIL import ImageChops
from pypdf import PdfReader
import pypdfium2 as pdfium


def check(output: Path) -> dict:
    found = sorted(output.glob("Fig*.pdf"), key=lambda p: int(re.search(r"\d+", p.name).group()))
    expected = [f"Fig{i}.pdf" for i in range(1, 5)]
    if [p.name for p in found] != expected:
        raise AssertionError(f"Expected {expected}; got {[p.name for p in found]}")
    results = []
    for index, path in enumerate(found, 1):
        reader = PdfReader(path)
        assert len(reader.pages) == 1, (path.name, "must contain exactly one page")
        page = reader.pages[0]
        width_mm = float(page.mediabox.width) * 25.4 / 72
        height_mm = float(page.mediabox.height) * 25.4 / 72
        expected_width = 165.1 if index == 1 else 82.55
        assert abs(width_mm - expected_width) < 0.05, (path.name, width_mm, expected_width)
        assert height_mm > 0, path.name
        text = page.extract_text() or ""
        assert "Force" in text or "Time" in text, (path.name, "native vector text is missing")
        assert not re.search(r"Figure\s+\d", text), (path.name, "caption was not excluded")
        document = pdfium.PdfDocument(path)
        bitmap = document[0].render(scale=3)
        image = bitmap.to_pil().convert("RGB")
        white = image.copy()
        white.paste("white", (0, 0, image.width, image.height))
        diff = ImageChops.difference(image, white).convert("L")
        bbox = diff.point(lambda value: 255 if value > 10 else 0).getbbox()
        assert bbox is not None, (path.name, "blank rendering")
        assert bbox[0] > 0 and bbox[1] > 0 and bbox[2] < image.width and bbox[3] < image.height, (
            path.name, "ink touches page boundary; crop may clip art", bbox, image.size)
        results.append({"file": path.name, "width_mm": round(width_mm, 4),
                        "height_mm": round(height_mm, 4), "searchable_text": True,
                        "ink_bbox_px": list(bbox), "render_size_px": list(image.size)})
        bitmap.close()
        document.close()
    return {"status": "pass", "pdfs": results}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    print(json.dumps(check(args.output), indent=2))
