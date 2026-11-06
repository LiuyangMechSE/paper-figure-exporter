"""File-size control: lossless first, never below the PPI floor, never JPEG."""
from __future__ import annotations

from io import BytesIO
from pathlib import Path
import sys
import tempfile
import unittest

import numpy as np
from PIL import Image, ImageChops, ImageStat
import pypdfium2 as pdfium
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (BooleanObject, DecodedStreamObject, DictionaryObject, EncodedStreamObject,
                           NameObject, NumberObject)

ENGINE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ENGINE))
from image_quality import assess_pdf  # noqa: E402
from size_control import fit_file_size  # noqa: E402

PAGE_W, PAGE_H = 468.0, 300.0  # 165.1 mm wide


def photo(w, h, seed=0):
    rng = np.random.default_rng(seed)
    y, x = np.mgrid[0:h, 0:w]
    rgb = np.stack([128 + 90*np.sin(x/37.0)*np.cos(y/29.0), 128 + 70*np.cos((x+y)/61.0), x*255//w], -1)
    return Image.fromarray((rgb + rng.normal(0, 4, (h, w, 3))).clip(0, 255).astype('uint8'), 'RGB')


def image_obj(writer, im, alpha=None, jpeg=False):
    if jpeg:
        stream = EncodedStreamObject()
        buf = BytesIO(); im.save(buf, 'JPEG', quality=92)
        stream._data = buf.getvalue()
        stream[NameObject('/Filter')] = NameObject('/DCTDecode')
    else:
        raw = DecodedStreamObject(); raw.set_data(im.tobytes())
        stream = raw.flate_encode()  # plain deflate, no predictor (old encoding)
    stream.update({NameObject('/Type'): NameObject('/XObject'), NameObject('/Subtype'): NameObject('/Image'),
                   NameObject('/Width'): NumberObject(im.width), NameObject('/Height'): NumberObject(im.height),
                   NameObject('/BitsPerComponent'): NumberObject(8),
                   NameObject('/ColorSpace'): NameObject('/DeviceRGB' if im.mode == 'RGB' else '/DeviceGray'),
                   NameObject('/Interpolate'): BooleanObject(False)})
    if alpha is not None:
        m = DecodedStreamObject(); m.set_data(alpha.tobytes()); m = m.flate_encode()
        m.update({NameObject('/Type'): NameObject('/XObject'), NameObject('/Subtype'): NameObject('/Image'),
                  NameObject('/Width'): NumberObject(alpha.width), NameObject('/Height'): NumberObject(alpha.height),
                  NameObject('/BitsPerComponent'): NumberObject(8), NameObject('/ColorSpace'): NameObject('/DeviceGray')})
        stream[NameObject('/SMask')] = writer._add_object(m)
    return writer._add_object(stream)


def make_pdf(path, big=(2600, 1300)):
    w = PdfWriter()
    page = w.add_blank_page(PAGE_W, PAGE_H)
    big_im = photo(*big)
    yy, xx = np.mgrid[0:big[1], 0:big[0]]
    alpha = Image.fromarray(np.where((xx - big[0]/2)**2/(big[0]/2)**2 + (yy - big[1]/2)**2/(big[1]/2)**2 < 1, 255, 0).astype('uint8'), 'L')
    line = Image.new('L', (300, 200), 255)
    for i in range(0, 300, 10):
        line.paste(0, (i, 0, i+2, 200))
    xobjects = DictionaryObject({
        NameObject('/Big'): image_obj(w, big_im, alpha),               # 200 pt wide -> 936 PPI
        NameObject('/Jpg'): image_obj(w, photo(900, 600, 1), jpeg=True),  # 150 pt wide -> 432 PPI
        NameObject('/Line'): image_obj(w, line)})                      # 60 pt wide -> 360 PPI
    font = DictionaryObject({NameObject('/Type'): NameObject('/Font'), NameObject('/Subtype'): NameObject('/Type1'),
                             NameObject('/BaseFont'): NameObject('/Helvetica')})
    page[NameObject('/Resources')] = DictionaryObject({NameObject('/XObject'): xobjects,
                                                      NameObject('/Font'): DictionaryObject({NameObject('/F1'): w._add_object(font)})})
    content = DecodedStreamObject()
    content.set_data(b'q 200 0 0 100 10 150 cm /Big Do Q q 150 0 0 100 250 150 cm /Jpg Do Q '
                     b'q 60 0 0 40 10 20 cm /Line Do Q 2 w 0 0 1 RG 100 40 m 400 120 l S '
                     b'BT /F1 12 Tf 120 20 Td (Vector label 123) Tj ET')
    page[NameObject('/Contents')] = w._add_object(content)
    with open(path, 'wb') as f:
        w.write(f)


def render(path, dpi=96):
    doc = pdfium.PdfDocument(str(path))
    im = doc[0].render(scale=dpi/72).to_pil().convert('RGB')
    doc.close()
    return im


class SizeControlTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix='size-control-')
        self.root = Path(self.tmp.name)
        self.pdf = self.root / 'Fig2.pdf'
        make_pdf(self.pdf)
        self.original = self.pdf.read_bytes()

    def tearDown(self):
        self.tmp.cleanup()

    def common_checks(self, report, min_ppi=600):
        reader = PdfReader(str(self.pdf))
        page = reader.pages[0]
        self.assertEqual([float(v) for v in page.mediabox], [0, 0, PAGE_W, PAGE_H], 'page size must not change')
        self.assertIn('Vector label 123', page.extract_text())
        filters = [str(o.get_object().get('/Filter')) for o in page['/Resources']['/XObject'].values()]
        self.assertEqual(filters.count('/DCTDecode'), 1, 'only the original JPEG may stay JPEG; nothing new is JPEG')
        q = assess_pdf(self.pdf, target_ppi=min_ppi)
        before = {round(r['min_ppi']) for r in assess_pdf(self.root / 'lossless' / 'Fig2.pdf', target_ppi=min_ppi)['images']} \
            if (self.root / 'lossless' / 'Fig2.pdf').exists() else set()
        for rec in q['images']:
            # Pictures above the floor stay at or above it; pictures already below are untouched.
            if rec['image_width_px'] in (900, 300):
                continue
            self.assertGreaterEqual(rec['min_ppi'] + 1e-6, min_ppi)
        diff = ImageStat.Stat(ImageChops.difference(render(self.pdf), render(self.root / 'lossless' / 'Fig2.pdf')).convert('L')).mean[0]
        self.assertLess(diff, 2.0, 'appearance must stay the same at normal viewing size')
        self.assertEqual((self.root / 'lossless' / 'Fig2.pdf').read_bytes(), self.original, 'lossless copy is the untouched input')
        return before

    def test_under_limit_is_untouched(self):
        report = fit_file_size(self.pdf, len(self.original) + 1, 600, self.root / 'lossless' / 'Fig2.pdf')
        self.assertFalse(report['changed'])
        self.assertEqual(self.pdf.read_bytes(), self.original)
        self.assertFalse((self.root / 'lossless').exists())

    def test_lossless_reencoding_alone_keeps_every_pixel(self):
        report = fit_file_size(self.pdf, int(len(self.original) * 0.9), 600, self.root / 'lossless' / 'Fig2.pdf')
        self.assertTrue(report['meets_limit'], report)
        self.assertEqual(report['method'], 'lossless')
        self.assertTrue(all(not p['resampled'] for p in report['pictures']))
        self.common_checks(report)
        # Exact samples: decoded big image and its mask are identical to the input.
        get = lambda path: {str(k): v.get_object() for k, v in PdfReader(str(path)).pages[0]['/Resources']['/XObject'].items()}
        a, b = get(self.root / 'lossless' / 'Fig2.pdf')['/Big'], get(self.pdf)['/Big']
        self.assertEqual(a.get_data(), b.get_data())
        self.assertEqual(a['/SMask'].get_object().get_data(), b['/SMask'].get_object().get_data())
        self.assertIn('/DecodeParms', b)

    def test_reduction_respects_ppi_floor_and_reaches_limit(self):
        limit = int(len(self.original) * 0.45)
        report = fit_file_size(self.pdf, limit, 600, self.root / 'lossless' / 'Fig2.pdf')
        self.assertTrue(report['meets_limit'], report)
        self.assertLessEqual(self.pdf.stat().st_size, limit)
        self.assertEqual(report['method'], 'lossless_plus_reduced_to_ppi_cap')
        self.assertGreaterEqual(report['ppi_cap'], 600)
        reduced = [p for p in report['pictures'] if p['resampled']]
        self.assertEqual(len(reduced), 1, 'only the picture above the floor may be reduced')
        self.assertGreaterEqual(reduced[0]['ppi_after'], 600)
        self.common_checks(report)

    def test_impossible_limit_stops_at_floor_and_explains(self):
        report = fit_file_size(self.pdf, 20_000, 600, self.root / 'lossless' / 'Fig2.pdf')
        self.assertFalse(report['meets_limit'])
        self.assertTrue(report['changed'])
        self.assertEqual(report['ppi_cap'], 600)
        self.assertTrue(report['notes'])
        self.common_checks(report)


if __name__ == '__main__':
    unittest.main()
