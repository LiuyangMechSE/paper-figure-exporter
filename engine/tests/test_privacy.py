"""Exported figures and reports must not carry personal information:
no folder paths or user names, no document/photo metadata, no hyperlinks."""
from __future__ import annotations

from io import BytesIO
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zlib

import numpy as np
from PIL import Image
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (ArrayObject, DecodedStreamObject, DictionaryObject, EncodedStreamObject, FloatObject,
                           NameObject, NumberObject, TextStringObject)

ENGINE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ENGINE))
from privacy import Redactor, strip_jpeg_metadata  # noqa: E402

spec = importlib.util.spec_from_file_location('export_figures_privacy', ENGINE / 'export_figures.py')
exporter = importlib.util.module_from_spec(spec)
spec.loader.exec_module(exporter)

SECRET_NAME = 'Jane Q. Researcher'
SECRET_PATH = r'C:\Users\jqresearcher\Secret Project\deck.pptx'


def jpeg_with_metadata(rgb):
    exif = Image.Exif()
    exif[0x013B] = SECRET_NAME            # Artist
    exif[0x010F] = 'SecretCam'            # Make
    exif[0x8825] = {2: (40.0, 6.0, 0.0)}  # GPS latitude
    buf = BytesIO()
    rgb.save(buf, 'JPEG', quality=92, exif=exif, comment=f'by {SECRET_NAME}'.encode())
    return buf.getvalue()


def photo(w=320, h=200):
    yy, xx = np.mgrid[0:h, 0:w]
    return Image.fromarray(np.stack([xx * 255 // w, yy * 255 // h, (xx + yy) % 256], -1).astype('uint8'), 'RGB')


def all_text(pdf_path):
    """Raw file bytes plus every decompressed stream, for searching."""
    raw = Path(pdf_path).read_bytes()
    chunks = [raw]
    reader = PdfReader(str(pdf_path))
    ids = set(reader.xref_objStm) | {i for table in reader.xref.values() for i in table}
    for i in sorted(ids):
        try:
            obj = reader.get_object(i)
        except Exception:
            continue
        if hasattr(obj, 'get_data'):
            try:
                chunks.append(obj.get_data())
            except Exception:
                pass
    return b'\n'.join(chunks)


class JpegMetadataTests(unittest.TestCase):
    def test_exif_gps_and_comment_removed_pixels_identical(self):
        data = jpeg_with_metadata(photo())
        self.assertIn(SECRET_NAME.encode(), data)
        clean = strip_jpeg_metadata(data)
        self.assertNotIn(SECRET_NAME.encode(), clean)
        self.assertNotIn(b'SecretCam', clean)
        self.assertNotIn(b'Exif\x00\x00', clean)
        self.assertLess(len(clean), len(data))
        a, b = Image.open(BytesIO(data)), Image.open(BytesIO(clean))
        self.assertEqual(a.tobytes(), b.tobytes(), 'decoded pixels must be identical')
        self.assertEqual(b.getexif(), {}, 'no EXIF left')

    def test_colour_profile_is_kept_and_garbage_is_untouched(self):
        buf = BytesIO()
        icc = b'\x00' * 128 + b'fake-profile-bytes'
        photo().save(buf, 'JPEG', icc_profile=icc)
        clean = strip_jpeg_metadata(buf.getvalue())
        self.assertEqual(Image.open(BytesIO(clean)).info.get('icc_profile'), icc)
        self.assertEqual(strip_jpeg_metadata(b'not a jpeg'), b'not a jpeg')


class RedactorTests(unittest.TestCase):
    def test_paths_replaced_in_any_case_or_slash_style(self):
        r = Redactor([(r'C:\Users\jqresearcher\Papers\Deck_figures_1', '<results>'),
                      (r'C:\Users\jqresearcher\Papers', '<deck folder>'),
                      (r'C:\Users\jqresearcher', '~'), ('C:', 'never')])
        value = {'source': r'c:\users\JQRESEARCHER\Papers\deck.pptx',
                 'list': ['C:/Users/jqresearcher/Papers/Deck_figures_1/Fig1.pdf', r'C:\Users\jqresearcher\AppData\x'],
                 'n': 3}
        out = r(value)
        self.assertEqual(out['source'], r'<deck folder>\deck.pptx')
        self.assertEqual(out['list'], ['<results>/Fig1.pdf', r'~\AppData\x'])
        self.assertEqual(out['n'], 3)
        self.assertNotIn('jqresearcher', json.dumps(out).lower())
        self.assertEqual(value['source'], r'c:\users\JQRESEARCHER\Papers\deck.pptx', 'input not modified')


class FigurePdfCleanupTests(unittest.TestCase):
    def test_crop_removes_metadata_links_and_photo_exif(self):
        from types import SimpleNamespace
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            src = td / 'native.pdf'
            w = PdfWriter()
            page = w.add_blank_page(400, 300)
            jpeg = EncodedStreamObject()
            jpeg._data = jpeg_with_metadata(photo())
            xmp = DecodedStreamObject()
            xmp.set_data(f'<x:xmpmeta><dc:creator>{SECRET_NAME}</dc:creator></x:xmpmeta>'.encode())
            xmp.update({NameObject('/Type'): NameObject('/Metadata'), NameObject('/Subtype'): NameObject('/XML')})
            jpeg.update({NameObject('/Type'): NameObject('/XObject'), NameObject('/Subtype'): NameObject('/Image'),
                         NameObject('/Width'): NumberObject(320), NameObject('/Height'): NumberObject(200),
                         NameObject('/BitsPerComponent'): NumberObject(8), NameObject('/ColorSpace'): NameObject('/DeviceRGB'),
                         NameObject('/Filter'): NameObject('/DCTDecode'), NameObject('/Metadata'): w._add_object(xmp)})
            link = DictionaryObject({NameObject('/Type'): NameObject('/Annot'), NameObject('/Subtype'): NameObject('/Link'),
                                     NameObject('/Rect'): ArrayObject([FloatObject(v) for v in (50, 50, 150, 100)]),
                                     NameObject('/A'): DictionaryObject({NameObject('/S'): NameObject('/Launch'),
                                                                         NameObject('/F'): TextStringObject(SECRET_PATH)})})
            page[NameObject('/Annots')] = ArrayObject([w._add_object(link)])
            page[NameObject('/PieceInfo')] = DictionaryObject({NameObject('/PowerPoint'): TextStringObject(SECRET_NAME)})
            page[NameObject('/Resources')] = DictionaryObject({NameObject('/XObject'): DictionaryObject({NameObject('/Im0'): w._add_object(jpeg)})})
            content = DecodedStreamObject(); content.set_data(b'q 160 0 0 100 50 50 cm /Im0 Do Q')
            page[NameObject('/Contents')] = w._add_object(content)
            w.add_metadata({'/Author': SECRET_NAME, '/Title': 'Secret Project'})
            with open(src, 'wb') as f:
                w.write(f)
            self.assertIn(SECRET_NAME.encode(), all_text(src))

            out = td / 'Fig1.pdf'
            c = dict(id='s1f1', width='full', bbox_pt=[0, 0, 400, 300], canvas_pt=[0, 0, 400, 300], warnings=[])
            exporter.crop_pdf(src, out, c, SimpleNamespace(trim_dpi=216, padding_pt=2, full_width_mm=None,
                                                          half_width_mm=None, preview_dpi=72))
            text = all_text(out)
            for secret in (SECRET_NAME.encode(), b'SecretCam', b'jqresearcher', b'Secret Project', b'Exif\x00\x00'):
                self.assertNotIn(secret, text, secret)
            reader = PdfReader(str(out))
            self.assertNotIn('/Annots', reader.pages[0])
            self.assertEqual(set(reader.metadata or {}) - {'/Title', '/Creator', '/Producer'}, set())
            image = reader.pages[0]['/Resources']['/XObject']['/Im0'].get_object()
            self.assertEqual(Image.open(BytesIO(image._data)).tobytes(),
                             Image.open(BytesIO(jpeg_with_metadata(photo()))).tobytes(), 'photo pixels identical')


class EndToEndPrivacyTests(unittest.TestCase):
    """Full run with a stand-in renderer: no folder path anywhere in the results."""

    def test_results_contain_no_folder_paths(self):
        with tempfile.TemporaryDirectory(prefix='private-user-folder-') as td:
            deck_dir = Path(td) / 'My Secret Project'
            deck_dir.mkdir()
            deck = deck_dir / 'Deck.pptx'
            deck.write_bytes((ENGINE / 'tests' / 'synthetic.pptx').read_bytes())
            out = deck_dir / 'Deck_figures_run'

            def fake_run(cmd, check, timeout):
                plan = json.loads(Path(cmd[cmd.index('-Plan') + 1]).read_text())
                rendered = Path(cmd[cmd.index('-OutputDir') + 1]); rendered.mkdir()
                from reportlab.pdfgen import canvas
                for f in plan['figures']:
                    x0, y0, x1, y1 = f['canvas_pt']
                    pdf = canvas.Canvas(str(rendered / (f['id'] + '.pdf')), pagesize=(x1 - x0, y1 - y0))
                    pdf.setAuthor(SECRET_NAME); pdf.setTitle('Secret Project')
                    pdf.rect(100, 200, 300, 120, fill=1); pdf.drawString(100, 180, 'Axis label'); pdf.save()
                for slide in plan['known_shapes']:
                    Image.new('RGB', (1280, 960), 'white').save(rendered / f'slide-{int(slide):03}.png')
                (rendered / 'render_status.json').write_text(json.dumps(
                    {'status': 'complete', 'source': str(deck), 'shape_inventory': [],
                     'warnings': [f'Temporary copy at {Path.home()}/x and {deck_dir}\\temp']}), encoding='utf-8')
                return subprocess.CompletedProcess(cmd, 0)

            original = exporter.subprocess.run
            exporter.subprocess.run = fake_run
            try:
                shell = Path(td) / 'fake_shell'; shell.write_text('')
                exporter.main([str(deck), '--output', str(out), '--full-width-mm', '165.1', '--offslide', 'clip',
                               '--powershell-exe', str(shell), '--preview-dpi', '72', '--max-file-mb', '5'])
            finally:
                exporter.subprocess.run = original

            self.assertTrue(list(out.glob('Fig*.pdf')), 'figures were exported')
            self.assertFalse(list((out / 'rendered').glob('*.pdf')), "PowerPoint's raw PDFs removed")
            self.assertTrue(list((out / 'rendered').glob('slide-*.png')), 'slide images kept for review.html')
            manifest = json.loads((out / 'manifest.json').read_text(encoding='utf-8'))
            self.assertEqual(manifest['source'], 'Deck.pptx')
            forbidden = [str(deck_dir), str(Path(td)), str(Path.home()), 'My Secret Project', SECRET_NAME, 'private-user-folder']
            for f in out.rglob('*'):
                if f.is_file():
                    data = all_text(f) if f.suffix == '.pdf' else f.read_bytes()
                    text = data.decode('utf-8', 'ignore').replace('\\\\', '\\')
                    for secret in forbidden:
                        self.assertNotIn(secret, text, f'{secret!r} found in {f.relative_to(out)}')


if __name__ == '__main__':
    unittest.main()
