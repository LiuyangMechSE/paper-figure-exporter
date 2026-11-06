"""Source-pixel, alpha, rotation, PDF-number and safety regressions.

Native inputs are constructed explicitly, without PowerPoint or production
mapping helpers. Rendering an independently built full-source reference proves
appearance and placement; larger image dimensions alone do not pass the test.
"""
from __future__ import annotations

import json
from pathlib import Path
import sys
import tempfile
import unittest

import numpy as np
from PIL import Image
import pypdfium2 as pdfium
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (ArrayObject, ContentStream, DecodedStreamObject,
                           DictionaryObject, NameObject, NumberObject)
from pptx import Presentation
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from preserve_images import restore_images


def native_pdf(path, rgba, matrix):
    """Minimal native-like PDF, independent of preservation implementation."""
    writer = PdfWriter()
    page = writer.add_blank_page(864, 648)
    image = DecodedStreamObject()
    image.set_data(rgba.convert('RGB').tobytes())
    image.update({NameObject('/Type'): NameObject('/XObject'),
                  NameObject('/Subtype'): NameObject('/Image'),
                  NameObject('/Width'): NumberObject(rgba.width),
                  NameObject('/Height'): NumberObject(rgba.height),
                  NameObject('/BitsPerComponent'): NumberObject(8),
                  NameObject('/ColorSpace'): NameObject('/DeviceRGB')})
    alpha = rgba.getchannel('A')
    if alpha.getextrema() != (255, 255):
        mask = DecodedStreamObject()
        mask.set_data(alpha.tobytes())
        mask.update({NameObject('/Type'): NameObject('/XObject'),
                     NameObject('/Subtype'): NameObject('/Image'),
                     NameObject('/Width'): NumberObject(rgba.width),
                     NameObject('/Height'): NumberObject(rgba.height),
                     NameObject('/BitsPerComponent'): NumberObject(8),
                     NameObject('/ColorSpace'): NameObject('/DeviceGray')})
        image[NameObject('/SMask')] = writer._add_object(mask.flate_encode())
    page[NameObject('/Resources')] = DictionaryObject({NameObject('/XObject'):
        DictionaryObject({NameObject('/Native'): writer._add_object(image.flate_encode())})})
    stream = DecodedStreamObject()
    # Cyan background + overlapping red vector bar test alpha and paint order.
    numbers = ' '.join(f'{n:.12f}' for n in matrix)
    stream.set_data(('0.1 0.6 0.8 rg 20 20 824 608 re f\n'
                     'q '+numbers+' cm /Native Do Q\n'
                     '1 0 0 rg 140 450 200 3 re f\n').encode())
    page[NameObject('/Contents')] = writer._add_object(stream)
    with open(path, 'wb') as f:
        writer.write(f)


def rendered(path):
    doc = pdfium.PdfDocument(str(path))
    page = doc[0]
    bitmap = page.render(scale=2)
    result = np.asarray(bitmap.to_pil().convert('RGB')).copy()
    bitmap.close(); page.close(); doc.close()
    return result


class PreserveImagesTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='source-image-qa-')
        self.root = Path(self.temp.name)
        # Rich colors + asymmetric corners distinguish rotations/flips.
        array = np.zeros((420, 600, 4), dtype=np.uint8)
        array[:, :] = [230, 245, 255, 255]
        array[:210, 300:] = [255, 200, 190, 255]
        array[210:, :300] = [190, 255, 170, 255]
        array[210:, 300:] = [240, 180, 255, 255]
        array[::20, :, :3] = 10
        array[:, ::20, :3] = 10
        self.image = Image.fromarray(array)

    def tearDown(self):
        self.temp.cleanup()

    def source(self, image, rotation=0, flip_h=False):
        asset = self.root / 'source.png'; image.save(asset)
        prs = Presentation(); prs.slide_width = Inches(12); prs.slide_height = Inches(9)
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        pic = slide.shapes.add_picture(str(asset), Inches(1), Inches(1), Inches(3.5), Inches(2.45))
        pic.rotation = rotation
        if flip_h:
            pic._element.spPr.xfrm.set('flipH', '1')
        source = self.root / 'source.pptx'; prs.save(source)
        return source, pic.shape_id

    def run_case(self, image, matrix, rotation=0, flip_h=False):
        source, sid = self.source(image, rotation, flip_h)
        before, after, reference = [self.root / p for p in ('native.pdf', 'restored.pdf', 'reference.pdf')]
        native_pdf(before, image.resize((200, 140), Image.Resampling.LANCZOS), matrix)
        native_pdf(reference, image, matrix)
        report = restore_images(source, before, after, 1, [sid], [0, 0, 864, 648])
        self.assertEqual(report['restored_resources'], 1, json.dumps(report, indent=2))
        self.assertFalse(report['resampling'])
        page = PdfReader(after).pages[0]
        self.assertEqual(page.get_contents().get_data(), PdfReader(before).pages[0].get_contents().get_data())
        preserved = list(page.images)[0].image.convert('RGBA')
        self.assertEqual(preserved.size, image.size)
        self.assertEqual(preserved.tobytes(), image.tobytes(), 'Output must contain exact source samples')
        # Form matrix must be a valid six-number PDF cm, never 1e-07 notation.
        form = page['/Resources']['/XObject']['/Native']
        operations = ContentStream(form, page.pdf).operations
        self.assertEqual([op for _, op in operations], [b'q', b're', b'W', b'n', b'cm', b'Do', b'Q'])
        self.assertEqual(len(operations[4][0]), 6)
        diff = np.abs(rendered(after).astype(float) - rendered(reference).astype(float))
        self.assertLess(diff.mean(), .15, 'Image missing, transformed, masked or reordered')
        return report, after, source

    def test_exact_original_pixels_and_vector_zorder(self):
        self.run_case(self.image, [252, 0, 0, 176.4, 72, 399.6])

    def test_rotation_flip_and_small_rounding_have_valid_pdf_numbers(self):
        # Clockwise90 + flipH maps image x upward, image y rightward.
        self.run_case(self.image, [.000004, 252, 176.4, 0, 109.8, 361.8],
                      rotation=90, flip_h=True)

    def test_source_alpha_and_transparent_hole_preserved(self):
        array = np.asarray(self.image).copy()
        array[:, :, 3] = np.linspace(0, 255, 600).astype('uint8')[None, :]
        array[130:230, 220:330] = [255, 0, 255, 0]
        self.run_case(Image.fromarray(array), [252, 0, 0, 176.4, 72, 399.6])

    def test_refuse_overwrite_and_keep_original_files(self):
        report, output, source = self.run_case(self.image, [252, 0, 0, 176.4, 72, 399.6])
        original, exported = source.read_bytes(), output.read_bytes()
        with self.assertRaises(FileExistsError):
            restore_images(source, self.root / 'native.pdf', output, 1, [2], [0, 0, 864, 648])
        self.assertEqual(source.read_bytes(), original)
        self.assertEqual(output.read_bytes(), exported)

    def test_ambiguous_overlapping_sources_retain_native_and_report(self):
        source, sid = self.source(self.image)
        prs = Presentation(source)
        duplicate = prs.slides[0].shapes.add_picture(str(self.root / 'source.png'),
                    Inches(1), Inches(1), Inches(3.5), Inches(2.45))
        prs.save(source)
        before, after = self.root / 'native.pdf', self.root / 'restored.pdf'
        native_pdf(before, self.image.resize((200, 140)), [252, 0, 0, 176.4, 72, 399.6])
        report = restore_images(source, before, after, 1, [sid, duplicate.shape_id], [0, 0, 864, 648])
        self.assertEqual(report['restored_resources'], 0)
        self.assertTrue(any('Ambiguous' in warning for r in report['images'] for warning in r['warnings']))
        self.assertTrue(np.array_equal(rendered(before), rendered(after)))

    def test_picture_effect_keeps_native_appearance_with_explicit_warning(self):
        source, sid = self.source(self.image)
        prs = Presentation(source)
        effect = OxmlElement('a:lum'); effect.set('bright', '25000')
        prs.slides[0].shapes[0]._element.blipFill.blip.append(effect)
        prs.save(source)
        before, after = self.root / 'native.pdf', self.root / 'restored.pdf'
        native_pdf(before, self.image.resize((200, 140)), [252, 0, 0, 176.4, 72, 399.6])
        report = restore_images(source, before, after, 1, [sid], [0, 0, 864, 648])
        self.assertEqual(report['restored_resources'], 0)
        self.assertTrue(any('Picture effects: lum' in warning for r in report['images'] for warning in r['warnings']))
        self.assertTrue(np.array_equal(rendered(before), rendered(after)))

    def test_jpeg_uses_exact_original_codestream_without_recompression(self):
        asset = self.root / 'source.jpg'
        self.image.convert('RGB').save(asset, quality=87)
        original = asset.read_bytes()
        prs = Presentation(); prs.slide_width = Inches(12); prs.slide_height = Inches(9)
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        picture = slide.shapes.add_picture(str(asset), Inches(1), Inches(1), Inches(3.5), Inches(2.45))
        source = self.root / 'source.pptx'; prs.save(source)
        before, after = self.root / 'native.pdf', self.root / 'restored.pdf'
        image = Image.open(asset).convert('RGBA')
        native_pdf(before, image.resize((200, 140), Image.Resampling.LANCZOS), [252, 0, 0, 176.4, 72, 399.6])
        report = restore_images(source, before, after, 1, [picture.shape_id], [0, 0, 864, 648])
        self.assertEqual(report['restored_resources'], 1)
        restored = list(PdfReader(after).pages[0].images)[0].indirect_reference.get_object()
        self.assertEqual(restored['/Filter'], '/DCTDecode')
        self.assertEqual(restored.get_data(), original)


if __name__ == '__main__':
    unittest.main(verbosity=2)


class FlattenedTransparencyTests(unittest.TestCase):
    """PowerPoint's PDF export can fill a picture's transparent areas (seen with
    transparent TIFFs, sometimes re-encoded as JPEG). The slide preview, which
    PowerPoint renders itself, decides whether the transparency is real."""

    SLIDE = (864.0, 648.0)            # 12 x 9 in
    BOX = (72.0, 72.0, 252.0, 176.4)  # picture left, top, width, height in points

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='flattened-alpha-')
        self.root = Path(self.temp.name)
        h, w = 420, 600
        yy, xx = np.mgrid[0:h, 0:w]
        rgba = np.zeros((h, w, 4), dtype=np.uint8)          # transparent pixels hide BLACK
        inside = ((xx - 260) / 210.0) ** 2 + ((yy - 210) / 170.0) ** 2 < 1
        rgba[..., 0] = (120 + 100 * np.sin(xx / 23.0)).astype(np.uint8)
        rgba[..., 1] = (140 + 80 * np.cos(yy / 31.0)).astype(np.uint8)
        rgba[..., 2] = 200
        rgba[..., 3] = np.where(inside, 255, 0)
        rgba[~inside, :3] = 0
        rgba[:30, 520:, :] = [250, 30, 30, 255]              # opaque asymmetric marker (detects flips)
        self.image = Image.fromarray(rgba, 'RGBA')

    def tearDown(self):
        self.temp.cleanup()

    def build(self, slide_shows_alpha=True, with_preview=True):
        asset = self.root / 'photo.tif'
        self.image.save(asset)
        prs = Presentation(); prs.slide_width = Inches(12); prs.slide_height = Inches(9)
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        left, top, width, height = self.BOX
        pic = slide.shapes.add_picture(str(asset), Inches(left/72), Inches(top/72), Inches(width/72), Inches(height/72))
        pic._element.spPr.xfrm.set('flipH', '1')
        source = self.root / 'deck.pptx'; prs.save(source)
        # Native PDF: flipped, downsampled to ~200 PPI and OPAQUE (black where transparent).
        flat = self.image.convert('RGB').resize((700, 490), Image.Resampling.LANCZOS)
        native = self.root / 'native.pdf'
        writer = PdfWriter(); page = writer.add_blank_page(*self.SLIDE)
        obj = DecodedStreamObject(); obj.set_data(flat.tobytes())
        obj.update({NameObject('/Type'): NameObject('/XObject'), NameObject('/Subtype'): NameObject('/Image'),
                    NameObject('/Width'): NumberObject(flat.width), NameObject('/Height'): NumberObject(flat.height),
                    NameObject('/BitsPerComponent'): NumberObject(8), NameObject('/ColorSpace'): NameObject('/DeviceRGB')})
        page[NameObject('/Resources')] = DictionaryObject({NameObject('/XObject'): DictionaryObject(
            {NameObject('/Im'): writer._add_object(obj.flate_encode())})})
        y0 = self.SLIDE[1] - top - height
        content = DecodedStreamObject()
        content.set_data(f'q {-width} 0 0 {height} {left + width} {y0} cm /Im Do Q'.encode())
        page[NameObject('/Contents')] = writer._add_object(content)
        with open(native, 'wb') as f:
            writer.write(f)
        # PowerPoint's own slide image (1280 px wide).
        preview = None
        if with_preview:
            k = 1280 / self.SLIDE[0]
            canvas = Image.new('RGB', (1280, round(self.SLIDE[1] * k)), 'white')
            shown = self.image if slide_shows_alpha else self.image.convert('RGB').convert('RGBA')
            shown = shown.transpose(Image.Transpose.FLIP_LEFT_RIGHT).resize((round(width*k), round(height*k)), Image.Resampling.LANCZOS)
            canvas.paste(shown, (round(left*k), round(top*k)), shown)
            preview = self.root / 'slide-001.png'; canvas.save(preview)
        return source, pic.shape_id, native, preview

    def restore(self, **kw):
        source, sid, native, preview = self.build(**kw)
        out = self.root / 'restored.pdf'
        report = restore_images(source, native, out, 1, [sid], [0, 0, *self.SLIDE],
                                slide_preview=preview, slide_size_pt=self.SLIDE)
        return report, out

    def test_transparency_shown_on_slide_is_restored(self):
        report, out = self.restore(slide_shows_alpha=True)
        rec = report['images'][0]
        self.assertEqual(rec['status'], 'restored_original_pixels', json.dumps(rec, indent=2))
        self.assertEqual(rec['slide_transparency_check']['verdict'], 'transparent')
        self.assertTrue(rec['notes'])
        restored = list(PdfReader(out).pages[0].images)[0].image.convert('RGBA')
        self.assertEqual(restored.tobytes(), self.image.tobytes(), 'exact source pixels AND alpha')
        # Rendered: the former black box is now white page, matching the slide.
        page = rendered(out)                      # scale 2: page[row, col] = page[y*2, x*2]
        left, top, width, height = self.BOX
        corner = page[int((top + height - 10) * 2), int((left + 10) * 2)]   # transparent bottom-left corner
        self.assertTrue((corner > 240).all(), f'former black fill must be gone: {corner}')
        # Flip preserved: the source's top-RIGHT red marker appears top-LEFT on the page.
        marker = page[int((top + 3) * 2), int((left + 10) * 2)]
        self.assertTrue(marker[0] > 200 and marker[1] < 80, marker)

    def test_fill_also_shown_on_slide_keeps_powerpoint_version(self):
        report, _ = self.restore(slide_shows_alpha=False)
        rec = report['images'][0]
        self.assertEqual(rec['status'], 'retained_native')
        self.assertEqual(rec['slide_transparency_check']['verdict'], 'opaque')

    def test_without_slide_image_nothing_is_guessed(self):
        report, _ = self.restore(with_preview=False)
        rec = report['images'][0]
        self.assertEqual(rec['status'], 'retained_native')
        self.assertIsNone(rec['slide_transparency_check']['verdict'])
