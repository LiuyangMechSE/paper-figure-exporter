"""Analyze-only regression and source/output safety tests; no PowerPoint needed."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
import tempfile
import unittest

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.util import Inches

MODULE_PATH = Path(__file__).resolve().parents[1] / "export_figures.py"
spec = importlib.util.spec_from_file_location("export_figures_review", MODULE_PATH)
exporter = importlib.util.module_from_spec(spec)
spec.loader.exec_module(exporter)


class ExporterReviewTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix="paper-figure-review-tests-")
        self.root = Path(self.tmp.name)
        self.prs = Presentation()
        self.prs.slide_width = Inches(10)
        self.prs.slide_height = Inches(7.5)
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])

    def tearDown(self):
        self.tmp.cleanup()

    def save(self):
        source = self.root / "fixture.pptx"
        self.prs.save(source)
        return source

    def analyze(self, config=None):
        source = self.save()
        args = exporter.parse_args([str(source), "--analyze-only", "--full-width-mm", "165.1"])
        return exporter.analyze(source, config or {}, args)

    def test_grouped_caption_does_not_discard_native_art(self):
        group = self.slide.shapes.add_group_shape()
        caption = group.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(.4))
        caption.text = "Figure 1 (full width). Caption inside grouped art."
        group.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(1), Inches(2), Inches(5), Inches(3))
        result = self.analyze()
        self.assertTrue(result["candidates"], "Grouped figure vanished as a caption role")
        self.assertIn(group.shape_id, [sid for c in result["candidates"] for sid in c["shape_ids"]])
        self.assertNotEqual(exporter.shape_info(group)["role"], "figure_label")

    def test_grouped_width_text_does_not_discard_native_art(self):
        group = self.slide.shapes.add_group_shape()
        title = group.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(.4))
        title.text = "full width"
        group.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(1), Inches(2), Inches(5), Inches(3))
        self.assertIsNone(exporter.shape_info(group)["role"])
        self.assertTrue(self.analyze()["candidates"])

    def test_mixed_native_table_and_vector_group_requires_review(self):
        table = self.slide.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(3), Inches(2))
        table.table.cell(0, 0).text = "Data"
        art = self.slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(5), Inches(1), Inches(3), Inches(2))
        group = self.slide.shapes.add_group_shape([table, art])
        result = self.analyze()
        candidates = [c for c in result["candidates"] if group.shape_id in c["shape_ids"]]
        self.assertEqual(len(candidates), 1)
        self.assertEqual(candidates[0]["kind"], "review", "Skipping whole group as table loses its vector artwork")

    def test_mixed_native_table_and_chart_group_requires_review(self):
        table = self.slide.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(3), Inches(2))
        data = CategoryChartData()
        data.categories = ["A", "B"]
        data.add_series("Value", [1, 2])
        chart = self.slide.shapes.add_chart(XL_CHART_TYPE.LINE, Inches(5), Inches(1),
                                            Inches(3), Inches(2), data)
        group = self.slide.shapes.add_group_shape([table, chart])
        candidates = [c for c in self.analyze()["candidates"] if group.shape_id in c["shape_ids"]]
        self.assertEqual(candidates[0]["kind"], "review")

    def test_hidden_table_caption_does_not_skip_visible_figure(self):
        self.slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(1), Inches(1), Inches(3), Inches(2))
        caption = self.slide.shapes.add_textbox(Inches(1), Inches(4), Inches(6), Inches(.4))
        caption.text = "Table 1. Hidden former caption"
        for node in caption.element.iter():
            if node.tag.endswith("}cNvPr"):
                node.set("hidden", "1")
        kinds = [c["kind"] for c in self.analyze()["candidates"]]
        self.assertNotIn("table", kinds)
        self.assertTrue(any(kind in ("figure", "review") for kind in kinds))

    def test_generated_overrides_preserve_manual_crop(self):
        art = self.slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(1), Inches(1), Inches(3), Inches(2))
        crop = [70, 70, 300, 220]
        result = self.analyze({"slides": {"1": {"groups": [{"shape_ids": [art.shape_id],
            "kind": "figure", "width": "half", "crop_pt": crop}]}}})
        out = self.root / "report"
        out.mkdir()
        exporter.write_report(out, result)
        generated = json.loads((out / "overrides.generated.json").read_text(encoding="utf-8"))
        self.assertEqual(generated["slides"]["1"]["groups"][0].get("crop_pt"), crop)

    def test_alternate_content_native_equation_is_selected(self):
        self.slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(1), Inches(1), Inches(3), Inches(2))
        equation = self.slide.shapes.add_textbox(Inches(4), Inches(1), Inches(3), Inches(1))
        equation.text = "E = m c squared"
        equation_id = equation.shape_id
        equation_xml = equation.element
        tree = equation_xml.getparent()
        position = tree.index(equation_xml)
        mc = "http://schemas.openxmlformats.org/markup-compatibility/2006"
        a14 = "http://schemas.microsoft.com/office/drawing/2010/main"
        alternate = etree.Element(f"{{{mc}}}AlternateContent", nsmap={"mc": mc, "a14": a14})
        choice = etree.SubElement(alternate, f"{{{mc}}}Choice", Requires="a14")
        tree.remove(equation_xml)
        choice.append(equation_xml)
        tree.insert(position, alternate)
        caption = self.slide.shapes.add_textbox(Inches(1), Inches(4), Inches(6), Inches(.4))
        caption.text = "Figure 1 (full width). Native equation must be included."
        manifest = self.analyze()
        inspected_ids = [s["id"] for s in manifest["slides"][0]["shapes"]]
        selected_ids = [sid for c in manifest["candidates"] if c["kind"] == "figure" for sid in c["shape_ids"]]
        self.assertIn(equation_id, inspected_ids, "AlternateContent equation was silently absent from inventory")
        self.assertIn(equation_id, selected_ids, "Native equation was silently absent from figure selection")
        self.assertEqual(manifest["slides"][0]["unsupported_objects"], [])

    def test_unsupported_top_level_object_is_reported_and_holds_figure(self):
        self.slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(1), Inches(1), Inches(3), Inches(2))
        caption = self.slide.shapes.add_textbox(Inches(1), Inches(4), Inches(6), Inches(.4))
        caption.text = "Figure 1 (full width). Unsupported object must not disappear."
        unknown = etree.Element("{http://schemas.openxmlformats.org/presentationml/2006/main}contentPart")
        self.slide.shapes._spTree.append(unknown)
        manifest = self.analyze()
        self.assertTrue(manifest["slides"][0]["unsupported_objects"])
        self.assertNotIn("figure", [c["kind"] for c in manifest["candidates"]])
        self.assertIn("review", [c["kind"] for c in manifest["candidates"]])

    def test_unknown_shape_override_is_rejected(self):
        self.slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(1), Inches(1), Inches(3), Inches(2))
        source = self.save()
        before = exporter.sha256(source)
        args = exporter.parse_args([str(source), "--analyze-only"])
        with self.assertRaisesRegex(ValueError, "unknown shape_ids"):
            exporter.analyze(source, {"slides": {"1": {"groups": [{"shape_ids": [9999]}]}}}, args)
        self.assertEqual(before, exporter.sha256(source))

    def test_duplicate_shape_assignment_is_rejected(self):
        art = self.slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(1), Inches(1), Inches(3), Inches(2))
        with self.assertRaisesRegex(ValueError, "duplicate shape assignment"):
            self.analyze({"slides": {"1": {"groups": [{"shape_ids": [art.shape_id]}, {"shape_ids": [art.shape_id]}]}}})

    def test_nonempty_output_is_not_overwritten(self):
        source = self.save()
        before = exporter.sha256(source)
        out = self.root / "output"
        out.mkdir()
        sentinel = out / "Fig1.pdf"
        sentinel.write_bytes(b"user-owned-output")
        with self.assertRaises(FileExistsError):
            exporter.main([str(source), "--output", str(out), "--analyze-only"])
        self.assertEqual(sentinel.read_bytes(), b"user-owned-output")
        self.assertEqual(before, exporter.sha256(source))

    def labelled_figure(self):
        label = self.slide.shapes.add_textbox(Inches(.1), Inches(.1), Inches(1), Inches(.3))
        label.text = "Figure 1"
        arts = [self.slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1 + 2.5*i), Inches(1), Inches(2), Inches(2))
                for i in range(3)]
        return arts

    def figure(self, manifest):
        figures = [c for c in manifest["candidates"] if c["label"] == "Figure 1"]
        self.assertEqual(len(figures), 1)
        return figures[0]

    def test_detached_guide_line_is_excluded_and_reported(self):
        arts = self.labelled_figure()
        guide = self.slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(.5), Inches(5), Inches(9.5), Inches(5))
        manifest = self.analyze()
        candidate = self.figure(manifest)
        self.assertEqual(sorted(candidate["shape_ids"]), sorted(a.shape_id for a in arts))
        self.assertEqual(candidate["kind"], "figure")
        unassigned = {u["id"]: u["reason"] for u in manifest["slides"][0]["unassigned_shapes"]}
        self.assertEqual(unassigned.get(guide.shape_id), "detached_guide_line")
        self.assertTrue(any("guide line" in w for w in manifest["slides"][0]["warnings"]))

    def test_lines_touching_or_inside_the_artwork_are_kept(self):
        arts = self.labelled_figure()
        arrow = self.slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(2), Inches(2), Inches(4), Inches(2))
        inside = self.slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(1), Inches(3.5), Inches(8), Inches(3.5))
        far_but_inside_bounds = self.slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(3.05), Inches(2), Inches(3.4), Inches(2))  # in the gap between ovals
        candidate = self.figure(self.analyze())
        for line in (arrow, far_but_inside_bounds):
            self.assertIn(line.shape_id, candidate["shape_ids"])
        # A detached line sticking 0.5 in below the art is a guide; keep_guide_lines retains it.
        self.assertNotIn(inside.shape_id, candidate["shape_ids"])
        source = self.save()
        args = exporter.parse_args([str(source), "--analyze-only", "--keep-guide-lines"])
        kept = self.figure(exporter.analyze(source, {}, args))
        self.assertIn(inside.shape_id, kept["shape_ids"])

    def test_offslide_parts_are_held_clipped_or_expanded_as_requested(self):
        self.labelled_figure()
        self.slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(-.3), Inches(1), Inches(1), Inches(1)).text = "a)"
        self.assertEqual(self.figure(self.analyze())["kind"], "review")
        source = self.save()
        slide_w, slide_h = 720.0, 540.0  # 10 x 7.5 in

        clip = self.figure(exporter.analyze(source, {}, exporter.parse_args([str(source), "--analyze-only", "--offslide", "clip"])))
        self.assertEqual(clip["kind"], "figure")
        self.assertEqual(clip["canvas_pt"], [0, 0, slide_w, slide_h], "Clip must render on the unchanged slide")
        self.assertGreaterEqual(clip["bbox_pt"][0], 0, "Width must be judged on the visible, clipped area")
        self.assertLess(clip["source_bbox_pt"][0], 0)
        self.assertTrue(any("clipped at the slide edge" in w for w in clip["warnings"]))

        expand = self.figure(exporter.analyze(source, {}, exporter.parse_args([str(source), "--analyze-only", "--offslide", "expand"])))
        self.assertEqual(expand["kind"], "figure")
        self.assertLess(expand["canvas_pt"][0], 0, "Expand must grow the canvas to keep the off-slide object")

    def test_content_entirely_off_slide_is_not_exported_blank(self):
        label = self.slide.shapes.add_textbox(Inches(.1), Inches(.1), Inches(1), Inches(.3))
        label.text = "Figure 1"
        for i in range(3):
            self.slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(11 + i), Inches(1), Inches(.5), Inches(.5))
        source = self.save()
        candidate = self.figure(exporter.analyze(source, {}, exporter.parse_args([str(source), "--analyze-only", "--offslide", "clip"])))
        self.assertEqual(candidate["kind"], "review")
        self.assertEqual(candidate["reason"], "content_entirely_off_slide")

    def test_edge_guide_overlapping_a_picture_box_is_excluded(self):
        # Real case: a full-height slide-edge guide whose x lies inside a plot picture's blank margin.
        arts = self.labelled_figure()                      # ovals spanning y 1-3 in, x 1-8 in
        plot = self.slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.5), Inches(1), Inches(1.6), Inches(1.5))
        guide = self.slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(8.08), Inches(.2), Inches(8.08), Inches(7.3))
        manifest = self.analyze()
        candidate = self.figure(manifest)
        self.assertNotIn(guide.shape_id, candidate["shape_ids"])
        self.assertIn(plot.shape_id, candidate["shape_ids"])
        unassigned = {u["id"]: u["reason"] for u in manifest["slides"][0]["unassigned_shapes"]}
        self.assertEqual(unassigned.get(guide.shape_id), "detached_guide_line")

    def test_line_overshooting_one_end_is_kept(self):
        self.labelled_figure()                             # ovals: y 1-3 in; first oval x 1-3 in
        # Touches the art and runs far past it at ONE end only (e.g. an extended axis): kept.
        axis = self.slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(2), Inches(1.2), Inches(2), Inches(5.5))
        self.assertIn(axis.shape_id, self.figure(self.analyze())["shape_ids"])

    def test_line_overshooting_both_ends_but_touching_a_line_is_kept(self):
        self.labelled_figure()
        cross = self.slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(5.75), Inches(.2), Inches(5.75), Inches(6.5))
        leader = self.slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(5.5), Inches(2), Inches(6), Inches(2))
        ids = self.figure(self.analyze())["shape_ids"]
        self.assertIn(cross.shape_id, ids)
        self.assertIn(leader.shape_id, ids)


if __name__ == "__main__":
    unittest.main(verbosity=2)
