"""Regression checks for vector crop, sizing, and reusable manual rectangles."""
import importlib.util
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

from pypdf import PdfReader
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('export_figures', ROOT / 'export_figures.py')
exporter = importlib.util.module_from_spec(spec)
spec.loader.exec_module(exporter)


class CropTests(unittest.TestCase):
    def test_manual_vector_crop_has_exact_size_and_preserves_input_rect(self):
        with tempfile.TemporaryDirectory() as td:
            source, output = Path(td)/'slide.pdf', Path(td)/'Fig1.pdf'
            pdf = canvas.Canvas(str(source), pagesize=(400, 300))
            pdf.rect(60, 100, 100, 30)
            pdf.drawString(65, 110, 'Force')
            pdf.save()
            c = dict(id='s1f1', width='half', crop_pt=[40,150,180,210], bbox_pt=[60,170,160,200], canvas_pt=[0,0,400,300], warnings=[])
            args = SimpleNamespace(trim_dpi=216,padding_pt=2,full_width_mm=165.1,half_width_mm=82.55)
            exporter.crop_pdf(source,output,c,args)
            page=PdfReader(output).pages[0]
            self.assertAlmostEqual(float(page.mediabox.width)*25.4/72,82.55,places=4)
            self.assertAlmostEqual(float(page.mediabox.width)/float(page.mediabox.height),140/60,places=5)
            self.assertIn('Force',page.extract_text())
            self.assertTrue(page['/Resources']['/Font'])
            self.assertEqual(c['crop_pt'],[40,150,180,210])
            self.assertEqual(c['crop_method'],'manual_rectangle')
            self.assertTrue(c['aspect_ratio_preserved'])

    def _slide(self, path, art_width):
        pdf = canvas.Canvas(str(path), pagesize=(720, 540))
        pdf.setFillColorRGB(0, 0, 0)
        pdf.rect(100, 200, art_width, 100, fill=1)
        pdf.setFont('Helvetica', 10)
        pdf.drawString(100, 180, 'Strain [mm/mm]')
        pdf.save()

    def _ink_bbox(self, path):
        import pypdfium2 as pdfium
        from PIL import Image, ImageChops
        doc = pdfium.PdfDocument(str(path)); im = doc[0].render(scale=4).to_pil().convert('RGB'); doc.close()
        box = ImageChops.difference(im, Image.new('RGB', im.size, 'white')).convert('L').point(lambda v: 255 if v > 6 else 0).getbbox()
        return [v / 4 for v in box], im.width / 4

    def test_half_width_match_full_uses_full_scale_and_centres_on_half_page(self):
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            full_src, half_src = td/'full.pdf', td/'half.pdf'
            self._slide(full_src, 500)   # wide figure
            self._slide(half_src, 200)   # narrow figure: would be enlarged to fill half width
            args = SimpleNamespace(trim_dpi=216, padding_pt=2, full_width_mm=165.1, half_width_mm=82.55,
                                   half_width_scale='match-full', preview_dpi=72)
            full = dict(id='s1f1', width='full', bbox_pt=[0,0,720,540], canvas_pt=[0,0,720,540], warnings=[])
            half = dict(id='s2f1', width='half', bbox_pt=[0,0,720,540], canvas_pt=[0,0,720,540], warnings=[])
            for c, src in ((full, full_src), (half, half_src)):
                exporter.measure_crop(src, c, args)
            ref = exporter.full_width_reference_scale([full, half], args)
            exporter.crop_pdf(full_src, td/'Fig1.pdf', full, args, reference_scale=ref, preview=False)
            exporter.crop_pdf(half_src, td/'Fig2.pdf', half, args, reference_scale=ref, preview=False)
            self.assertAlmostEqual(full['content_scale'], ref, places=9)
            self.assertAlmostEqual(half['content_scale'], ref, places=9, msg='text prints at the same size as in full-width figures')
            page = PdfReader(td/'Fig2.pdf').pages[0]
            self.assertAlmostEqual(float(page.mediabox.width)*25.4/72, 82.55, places=3, msg='page is exactly half width')
            self.assertIn('Strain', page.extract_text())
            (left, top, right, bottom), page_w = self._ink_bbox(td/'Fig2.pdf')
            self.assertGreater(left, 5, 'white side margin present')
            self.assertAlmostEqual(left, page_w - right, delta=1.0, msg='content centred')
            self.assertAlmostEqual(right - left, (200 + 0) * ref, delta=2.5, msg='content drawn at the reference scale')
            self.assertGreater(half['side_margin_mm'], 2)
            # Default 'fill' keeps the old behaviour: content fills the half width.
            fill_args = SimpleNamespace(**{**vars(args), 'half_width_scale': 'fill'})
            half2 = dict(half, warnings=[]); half2.pop('content_scale', None)
            exporter.measure_crop(half_src, half2, fill_args)
            exporter.crop_pdf(half_src, td/'Fig3.pdf', half2, fill_args, reference_scale=ref, preview=False)
            self.assertGreater(half2['content_scale'], ref)
            self.assertEqual(half2['side_margin_mm'], 0)

    def test_half_width_too_wide_at_reference_scale_is_reduced_to_fit(self):
        with tempfile.TemporaryDirectory() as td:
            td = Path(td); src = td/'half.pdf'
            self._slide(src, 400)
            args = SimpleNamespace(trim_dpi=216, padding_pt=2, full_width_mm=165.1, half_width_mm=82.55,
                                   half_width_scale='match-full', preview_dpi=72)
            c = dict(id='s1f1', width='half', bbox_pt=[0,0,720,540], canvas_pt=[0,0,720,540], warnings=[])
            exporter.crop_pdf(src, td/'Fig1.pdf', c, args, reference_scale=0.9, preview=False)
            page = PdfReader(td/'Fig1.pdf').pages[0]
            self.assertAlmostEqual(float(page.mediabox.width)*25.4/72, 82.55, places=3)
            self.assertLess(c['content_scale'], 0.9)
            self.assertTrue(any('reduced to fit' in w for w in c['warnings']))

    def test_blank_render_is_reported_and_does_not_create_pdf(self):
        with tempfile.TemporaryDirectory() as td:
            source, output = Path(td)/'blank.pdf', Path(td)/'Fig1.pdf'
            pdf=canvas.Canvas(str(source),pagesize=(400,300)); pdf.showPage(); pdf.save()
            c=dict(id='s1f1',width='full',bbox_pt=[10,10,20,20],warnings=[])
            args=SimpleNamespace(trim_dpi=216,padding_pt=2,full_width_mm=None,half_width_mm=None)
            with self.assertRaisesRegex(ValueError,'Blank rendering'):
                exporter.crop_pdf(source,output,c,args)
            self.assertFalse(output.exists())


if __name__=='__main__':
    unittest.main(verbosity=2)
