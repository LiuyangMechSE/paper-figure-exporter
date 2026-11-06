"""Measure placed samples, not nominal metadata or clipping-box dimensions."""
import sys
import tempfile
import unittest
from pathlib import Path

from pypdf import PdfWriter
from pypdf.generic import (ArrayObject, BooleanObject, DecodedStreamObject,
                           DictionaryObject, FloatObject, NameObject, NumberObject)

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from image_quality import assess_pdf


def dictionary(**items):
    return DictionaryObject({NameObject('/'+k): v for k, v in items.items()})


def numbers(items):
    return ArrayObject([FloatObject(v) for v in items])


class QualityTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.writer = PdfWriter()
        self.page = self.writer.add_blank_page(width=500, height=500)

    def image(self, width, height, gray=False, **extra):
        image = DecodedStreamObject()
        image.set_data(bytes([127])*(width*height*(1 if gray else 3)))
        image.update(dictionary(Type=NameObject('/XObject'), Subtype=NameObject('/Image'),
                                Width=NumberObject(width), Height=NumberObject(height),
                                BitsPerComponent=NumberObject(8),
                                ColorSpace=NameObject('/DeviceGray' if gray else '/DeviceRGB'),
                                **extra))
        return self.writer._add_object(image.flate_encode())

    def form(self, content, resources, matrix=(1, 0, 0, 1, 0, 0)):
        form = DecodedStreamObject()
        form.set_data(content.encode('ascii'))
        form.update(dictionary(Type=NameObject('/XObject'), Subtype=NameObject('/Form'),
                               BBox=numbers([0, 0, 200, 200]), Matrix=numbers(matrix),
                               Resources=resources))
        return self.writer._add_object(form)

    def assess(self, content, resources=None, **kwargs):
        stream = DecodedStreamObject()
        stream.set_data(content.encode('latin1'))
        self.page[NameObject('/Contents')] = self.writer._add_object(stream)
        self.page[NameObject('/Resources')] = resources or dictionary()
        path = Path(self.temp.name)/'quality.pdf'
        with path.open('wb') as handle:
            self.writer.write(handle)
        return assess_pdf(path, **kwargs)

    def test_rotation_uses_axis_lengths_and_pdf_false_is_false(self):
        im = self.image(1200, 600, ImageMask=BooleanObject(False), Interpolate=BooleanObject(False))
        result = self.assess('q 0 144 -72 0 200 20 cm /Im Do Q', dictionary(XObject=dictionary(Im=im)))
        self.assertEqual(result['image_count'], 1)
        image = result['images'][0]
        self.assertAlmostEqual(image['dpi_x'], 600)
        self.assertAlmostEqual(image['dpi_y'], 600)
        self.assertAlmostEqual(image['width_mm'], 50.8)
        self.assertFalse(image['interpolate'])
        self.assertFalse(image['stencil'])
        self.assertTrue(result['meets_target'])

    def test_nested_form_and_reused_image_have_separate_placement_ppi(self):
        im = self.image(600, 300)
        child = self.form('q 36 0 0 18 0 0 cm /Im Do Q', dictionary(XObject=dictionary(Im=im)), matrix=(2, 0, 0, 2, 0, 0))
        outer = self.form('/Child Do', dictionary(XObject=dictionary(Child=child)))
        result = self.assess('q /Outer Do Q q 2 0 0 2 120 0 cm /Outer Do Q', dictionary(XObject=dictionary(Outer=outer)))
        self.assertEqual(result['image_count'], 2)
        self.assertEqual(result['images'][0]['xref'], result['images'][1]['xref'])
        self.assertAlmostEqual(result['images'][0]['dpi_x'], 600)
        self.assertAlmostEqual(result['images'][1]['dpi_x'], 300)
        self.assertAlmostEqual(result['min_ppi'], 300)
        self.assertEqual(result['below_target_count'], 1)

    def test_mask_limits_quality_but_is_not_counted_as_another_picture(self):
        mask = self.image(300, 150, gray=True)
        im = self.image(1200, 600, SMask=mask)
        result = self.assess('q 144 0 0 72 0 0 cm /Im Do Q', dictionary(XObject=dictionary(Im=im)))
        self.assertEqual(result['image_count'], 1)
        image = result['images'][0]
        self.assertAlmostEqual(image['dpi_x'], 600)
        self.assertAlmostEqual(result['min_ppi'], 150)
        self.assertEqual(image['masks'][0]['required_width_px'], 1200)
        self.assertTrue(image['mask_below_target'])
        self.assertAlmostEqual(image['max_scale_at_target'], 0.25)

    def test_zero_opacity_and_off_page_images_are_not_target_failures(self):
        good = self.image(600, 600)
        low = self.image(10, 10)
        resources = dictionary(XObject=dictionary(Good=good, Low=low),
                               ExtGState=dictionary(Zero=dictionary(ca=FloatObject(0))))
        result = self.assess('q 72 0 0 72 0 0 cm /Good Do Q q /Zero gs 72 0 0 72 80 0 cm /Low Do Q q 72 0 0 72 900 0 cm /Low Do Q', resources)
        self.assertEqual(result['placement_count'], 3)
        self.assertEqual(result['image_count'], 1)
        self.assertTrue(result['meets_target'])
        self.assertFalse(result['images'][1]['potentially_visible'])

    def test_clipping_does_not_inflate_ppi_and_small_icons_still_fail(self):
        body = self.image(600, 600)
        icon = self.image(10, 20)
        result = self.assess('q 0 0 10 10 re W n 144 0 0 144 0 0 cm /Body Do Q q 2 0 0 4 20 20 cm /Icon Do Q', dictionary(XObject=dictionary(Body=body, Icon=icon)))
        self.assertAlmostEqual(result['images'][0]['dpi_x'], 300)
        self.assertAlmostEqual(result['images'][0]['visible_area_pt2_upper_bound'], 100)
        self.assertTrue(result['images'][1]['small_element'])
        self.assertEqual(result['below_target_count'], 2)
        self.assertFalse(result['contains_vector_paths'])  # A clip rectangle is not painted vector artwork.

    def test_userunit_and_remedies_use_physical_size(self):
        self.page[NameObject('/UserUnit')] = FloatObject(2)
        im = self.image(600, 300)
        result = self.assess('q 72 0 0 36 0 0 cm /Im Do Q', dictionary(XObject=dictionary(Im=im)))
        image = result['images'][0]
        self.assertAlmostEqual(image['dpi_x'], 300)
        self.assertEqual(image['required_width_px'], 1200)
        self.assertEqual(image['required_height_px'], 600)
        self.assertAlmostEqual(image['max_scale_at_target'], 0.5)
        self.assertAlmostEqual(image['max_width_mm_at_target'], 25.4)

    def test_shear_uses_worst_direction_not_axis_only(self):
        im = self.image(600, 600)
        result = self.assess('q 72 0 36 72 0 0 cm /Im Do Q', dictionary(XObject=dictionary(Im=im)))
        image = result['images'][0]
        self.assertAlmostEqual(image['dpi_x'], 600)
        self.assertLess(image['min_ppi'], image['dpi_y'])
        self.assertGreater(image['required_pixels_uniform_for_all_directions'][0], image['required_width_px'])

    def test_vector_and_native_text_have_no_raster_dpi(self):
        font = dictionary(Type=NameObject('/Font'), Subtype=NameObject('/Type1'), BaseFont=NameObject('/Helvetica'))
        result = self.assess('0 0 10 10 re f BT /F 12 Tf 10 10 Td (Native text) Tj ET', dictionary(Font=dictionary(F=font)))
        self.assertEqual(result['image_count'], 0)
        self.assertIsNone(result['min_ppi'])
        self.assertEqual(result['status'], 'no_raster_images')
        self.assertTrue(result['contains_native_text'])
        self.assertTrue(result['contains_vector_paths'])
        self.assertEqual(result['vector']['path_paint_ops'], 1)

    def test_external_transparency_mask_requires_review(self):
        im = self.image(600, 600)
        group = self.form('0 0 10 10 re f', dictionary())
        resources = dictionary(XObject=dictionary(Im=im), ExtGState=dictionary(G=dictionary(SMask=dictionary(S=NameObject('/Luminosity'), G=group))))
        result = self.assess('q /G gs 72 0 0 72 0 0 cm /Im Do Q', resources)
        self.assertFalse(result['assessment_complete'])
        self.assertFalse(result['meets_target'])
        self.assertEqual(result['status'], 'needs_review')
        self.assertTrue(result['issues'])
        self.assertIsNone(result['images'][0]['max_scale_at_target'])
        self.assertIsNone(result['max_figure_scale_at_target'])

    def test_inline_image_is_included(self):
        result = self.assess('q 0.12 0 0 0.12 0 0 cm BI /W 1 /H 1 /CS /RGB /BPC 8 ID \xff\x00\x00 EI Q')
        self.assertEqual(result['image_count'], 1)
        self.assertTrue(result['images'][0]['inline'])
        self.assertAlmostEqual(result['min_ppi'], 600)

    def test_tiling_pattern_content_is_not_silently_certified(self):
        pattern = dictionary(Type=NameObject('/Pattern'), PatternType=NumberObject(1))
        result = self.assess('/Pattern cs /P scn 0 0 10 10 re f', dictionary(Pattern=dictionary(P=pattern)))
        self.assertFalse(result['assessment_complete'])
        self.assertEqual(result['status'], 'needs_review')

    def test_type3_font_resources_require_review(self):
        font = dictionary(Type=NameObject('/Font'), Subtype=NameObject('/Type3'))
        result = self.assess('BT /F 12 Tf (Glyphs) Tj ET', dictionary(Font=dictionary(F=font)))
        self.assertFalse(result['assessment_complete'])
        self.assertEqual(result['status'], 'needs_review')

    def test_annotation_appearance_requires_review(self):
        appearance = self.form('0 0 10 10 re f', dictionary())
        annotation = dictionary(Type=NameObject('/Annot'), Subtype=NameObject('/Stamp'), AP=dictionary(N=appearance))
        self.page[NameObject('/Annots')] = ArrayObject([self.writer._add_object(annotation)])
        result = self.assess('0 0 10 10 re f')
        self.assertFalse(result['assessment_complete'])
        self.assertEqual(result['status'], 'needs_review')


if __name__ == '__main__':
    unittest.main(verbosity=2)
