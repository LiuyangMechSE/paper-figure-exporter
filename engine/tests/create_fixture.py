"""Create a deterministic small deck for exporter integration/visual tests.

This is wholly synthetic. It does not open or modify any user presentation.
Run: python tests/create_fixture.py
"""

from __future__ import annotations

import json
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from openpyxl import Workbook
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parent
BLUE = RGBColor(0, 130, 205)
ORANGE = RGBColor(230, 125, 40)
BLACK = RGBColor(0, 0, 0)


def text(shapes, content, x, y, w, h=0.35, size=14, name=None):
    box = shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    box.name = name or content
    box.text_frame.margin_left = 0
    box.text_frame.margin_top = 0
    box.text_frame.margin_right = 0
    box.text_frame.margin_bottom = 0
    p = box.text_frame.paragraphs[0]
    p.text = content
    p.font.name = "Arial"
    p.font.size = Pt(size)
    p.font.color.rgb = BLACK
    return box


def panel(shapes, x, y, w, h, color, label):
    rect = shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    rect.name = f"Panel {label} outline"
    rect.fill.background()
    rect.line.color.rgb = BLACK
    rect.line.width = Pt(0.8)
    text(shapes, label, x + 0.08, y + 0.03, 0.3, size=16)
    for i in range(5):
        point = shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.45 + i * (w - 0.9) / 4),
                                 Inches(y + h - 0.55 - (i / 4) ** 1.5 * (h - 0.9)),
                                 Inches(0.09), Inches(0.09))
        point.fill.solid()
        point.fill.fore_color.rgb = color
        point.line.fill.background()
        point.name = f"Panel {label} data marker {i}"
    text(shapes, "Time (s)", x + w / 2 - 0.5, y + h + 0.08, 1.5, size=12)
    text(shapes, "Force (N)", x + 0.25, y + 0.32, 1.5, size=12)


def diagram(slide, x, y, w, h, name):
    group = slide.shapes.add_group_shape()
    group.name = name
    gap = 0.35
    pw = (w - gap) / 2
    panel(group.shapes, x, y, pw, h, BLUE, "a")
    panel(group.shapes, x + pw + gap, y, pw, h, ORANGE, "b")
    return group


def raster_table():
    path = ROOT / "raster_table.png"
    image = Image.new("RGB", (1000, 430), "white")
    draw = ImageDraw.Draw(image)
    try:
        font = ImageFont.truetype("arial.ttf", 30)
    except OSError:
        font = ImageFont.load_default(size=30)
    rows = [["Sample", "Length (mm)", "Force (N)"],
            ["A", "100", "1.2"], ["B", "110", "1.4"], ["C", "120", "1.6"]]
    for row in range(5):
        draw.line((20, 20 + row * 95, 980, 20 + row * 95), fill="black", width=3)
    for x in (20, 330, 650, 980):
        draw.line((x, 20, x, 400), fill="black", width=3)
    for i, row in enumerate(rows):
        for j, value in enumerate(row):
            draw.text((40 + j * 320, 47 + i * 95), value, fill="black", font=font)
    image.save(path)
    return path


def build():
    ROOT.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    notes = []

    # A true group containing vector shapes and text, intended as one full-width figure.
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    diagram(slide, 1.0, 1.4, 8.0, 3.0, "Full width grouped figure")
    text(slide.shapes, "Figure 1 (full width). Synthetic two-panel force comparison.",
         1.0, 5.0, 8.2, size=14, name="Figure 1 caption")
    notes.append({"slide": 1, "expect": "figure", "width": "full", "pdf": "Fig1.pdf",
                  "check": "Keep both vector panels and text; exclude caption; crop around art."})

    # Native PowerPoint table with explicit table caption.
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    text(slide.shapes, "Table 1. Native table must be skipped.", 1, 0.8, 8, size=16)
    table = slide.shapes.add_table(4, 3, Inches(1), Inches(1.5), Inches(8), Inches(3)).table
    for r, row in enumerate([["Sample", "Length (mm)", "Force (N)"], ["A", "100", "1.2"],
                             ["B", "110", "1.4"], ["C", "120", "1.6"]]):
        for c, value in enumerate(row):
            table.cell(r, c).text = value
    notes.append({"slide": 2, "expect": "table", "pdf": None,
                  "check": "Skip native table and retain traceable reason in manifest."})

    # Native chart; PowerPoint's PDF renderer should preserve vector paths/fonts.
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    chart_data = CategoryChartData()
    chart_data.categories = [0, 1, 2, 3, 4]
    chart_data.add_series("Measured", [1, 1.2, 1.7, 2.2, 3.1])
    chart_data.add_series("Model", [1, 1.3, 1.6, 2.1, 3.0])
    chart = slide.shapes.add_chart(XL_CHART_TYPE.LINE, Inches(1), Inches(1.3), Inches(4),
                                   Inches(3.5), chart_data).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.has_title = False
    chart.category_axis.has_title = True
    chart.category_axis.axis_title.text_frame.text = "Time (s)"
    chart.value_axis.has_title = True
    chart.value_axis.axis_title.text_frame.text = "Force (N)"
    text(slide.shapes, "Figure 2 (half width). Native chart.", 1, 5.1, 5, size=14)
    notes.append({"slide": 3, "expect": "figure", "width": "half", "pdf": "Fig2.pdf",
                  "check": "Native line chart and legend remain intact; exclude caption."})

    # Two separately captioned figures on one slide, ordered left to right.
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    diagram(slide, 0.5, 1.5, 4.0, 2.8, "Left independent figure")
    diagram(slide, 5.5, 1.5, 4.0, 2.8, "Right independent figure")
    text(slide.shapes, "Figure 3 (half width). Left figure.", 0.5, 5, 4.0, size=13)
    text(slide.shapes, "Figure 4 (half width). Right figure.", 5.5, 5, 4.0, size=13)
    notes.append({"slide": 4, "expect": "two figures", "width": ["half", "half"],
                  "pdf": ["Fig3.pdf", "Fig4.pdf"],
                  "check": "Split at separate figure anchors; preserve panels inside each group."})

    # Raster tables cannot safely be classified from PPTX structural type alone.
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    picture = slide.shapes.add_picture(str(raster_table()), Inches(1), Inches(1.5), width=Inches(8))
    picture.name = "Unlabelled bitmap 1"
    notes.append({"slide": 5, "expect": "review", "pdf": None,
                  "check": "Unlabelled picture of table should be withheld, with visible preview."})

    # Cached Excel object display; the exporter must not activate or recalculate it.
    book = Workbook()
    ws = book.active
    for row in [["Sample", "Length (mm)", "Force (N)"], ["A", 100, 1.2], ["B", 110, 1.4]]:
        ws.append(row)
    object_file = ROOT / "embedded_table.xlsx"
    book.save(object_file)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_ole_object(str(object_file), "Excel.Sheet.12", Inches(1), Inches(1.5),
                               Inches(8), Inches(3.44), icon_file=str(ROOT / "raster_table.png"),
                               icon_width=Inches(8), icon_height=Inches(3.44))
    notes.append({"slide": 6, "expect": "table or review", "pdf": None,
                  "check": "Excel OLE is not assumed to be figure; disclose cached appearance."})

    # Edge cases deliberately follow the reliable figure sequence.
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide._element.set("show", "0")
    diagram(slide, 1, 1.5, 8, 3, "Hidden figure")
    text(slide.shapes, "Figure 5 (full width). Hidden draft.", 1, 5, 8, size=14)
    notes.append({"slide": 7, "expect": "skip unless include-hidden", "pdf": None,
                  "check": "Record hidden-slide skip; do not silently omit it."})

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    diagram(slide, 7, 1.5, 5, 3, "Off-slide figure")
    text(slide.shapes, "Figure 6 (full width). Intentionally extends past slide edge.",
         1, 5, 8, size=14)
    notes.append({"slide": 8, "expect": "review", "pdf": None,
                  "check": "Do not silently clip content extending beyond slide edge."})

    path = ROOT / "synthetic.pptx"
    prs.save(path)
    (ROOT / "expected.json").write_text(json.dumps({"slide_size_inches": [10, 7.5],
        "test_widths_mm": {"full": 165.1, "half": 82.55}, "slides": notes},
        indent=2), encoding="utf-8")
    print(path)
    print(f"Created {len(prs.slides)} slides. First four expected PDFs: Fig1-Fig4.")


if __name__ == "__main__":
    build()
