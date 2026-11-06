"""Independent PDF appearance/pixel audit for source-image restoration.

This uses no exporter code. Compare the native PowerPoint PDF against its
restored counterpart *before final cropping*. Geometry/drawing streams and
embedded fonts must be identical. Appearance differences are measured at a
common size and reported, while exact decoded source pixels prove that detail
came from a source file rather than an upsampled native PDF.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path

import numpy as np
from PIL import Image, ImageFilter
import pypdfium2 as pdfium
from pypdf import PdfReader


def sha(data):
    return hashlib.sha256(data).hexdigest()


def value(v):
    """Stable, indirect-object-number-independent PDF scalar serialization."""
    if hasattr(v, 'get_object'):
        v = v.get_object()
    if isinstance(v, dict):
        return {str(k): value(x) for k, x in sorted(v.items())
                if k not in ('/Length', '/Parent', '/P')}
    if isinstance(v, (list, tuple)):
        return [value(x) for x in v]
    if isinstance(v, bytes):
        return {'bytes': sha(v)}
    return str(v)


def inspect_pdf(path):
    reader = PdfReader(path)
    pages = []
    for page in reader.pages:
        streams, images, fonts, resources_used = {}, {}, {}, {}

        def walk(stream, resources, prefix='page', ancestors=()):
            streams[prefix] = sha(stream.get_data()) if stream else None
            resources = resources.get_object() if resources else {}
            for name, ref in resources.get('/Font', {}).items():
                font = ref.get_object()
                # Names, encodings and embedded program bytes survive rewriting.
                record = {'name': str(font.get('/BaseFont')),
                          'subtype': str(font.get('/Subtype')),
                          'encoding': value(font.get('/Encoding')),
                          'unicode': None, 'programs': []}
                if font.get('/ToUnicode'):
                    record['unicode'] = sha(font['/ToUnicode'].get_data())
                descendants = font.get('/DescendantFonts', [ref])
                for descendant in descendants:
                    descriptor = descendant.get_object().get('/FontDescriptor')
                    if descriptor:
                        descriptor = descriptor.get_object()
                        for key in ('/FontFile', '/FontFile2', '/FontFile3'):
                            if descriptor.get(key):
                                record['programs'].append(sha(descriptor[key].get_data()))
                fonts[prefix + str(name)] = record
            for name, ref in resources.get('/XObject', {}).items():
                obj = ref.get_object()
                key = prefix + str(name)
                resources_used[key] = str(obj.get('/Subtype'))
                if obj.get('/Subtype') == '/Image':
                    item = {'pixels': [int(obj['/Width']), int(obj['/Height'])],
                            'decoded_sha256': sha(obj.get_data()),
                            'bits': int(obj.get('/BitsPerComponent', 1)),
                            'colorspace': str(obj.get('/ColorSpace')),
                            'mask': value(obj.get('/Mask')), 'soft_mask': None}
                    if obj.get('/SMask'):
                        mask = obj['/SMask']
                        item['soft_mask'] = {
                            'pixels': [int(mask['/Width']), int(mask['/Height'])],
                            'decoded_sha256': sha(mask.get_data())}
                    images[key] = item
                elif obj.get('/Subtype') == '/Form':
                    identity = getattr(ref, 'idnum', id(obj))
                    if identity in ancestors:
                        raise AssertionError('Recursive form in QA input')
                    walk(obj, obj.get('/Resources', resources), key,
                         ancestors + (identity,))

        content = page.get_contents()
        walk(content, page.get('/Resources'))
        pages.append({'mediabox': [float(v) for v in page.mediabox],
                      'cropbox': [float(v) for v in page.cropbox],
                      'rotation': page.rotation, 'streams': streams,
                      'images': images, 'fonts': fonts,
                      'xobjects': resources_used,
                      'text': page.extract_text()})
    return pages


def source_matches(pdf_path, asset_dir):
    """Exact decoded RGB and alpha matching; no interpolated comparison."""
    assets = {}
    for path in Path(asset_dir).glob('*.png'):
        img = Image.open(path).convert('RGBA')
        assets[path.name] = img
    records = []
    for page_no, page in enumerate(PdfReader(pdf_path).pages, 1):
        for image in page.images:
            pixels = image.image.convert('RGBA')
            matches = []
            for name, asset in assets.items():
                if asset.size == pixels.size and asset.tobytes() == pixels.tobytes():
                    matches.append({'asset': name, 'type': 'exact_full_RGBA'})
                # Fixture's deliberate 20%,20%,10%,10% crop.
                if name == 'pattern.png':
                    crop = asset.crop((420, 294, 1890, 1323))
                    if crop.size == pixels.size and crop.tobytes() == pixels.tobytes():
                        matches.append({'asset': name, 'type': 'exact_visible_crop_RGBA'})
            records.append({'page': page_no, 'name': image.name,
                            'pixels': list(pixels.size), 'matches': matches})
    return records


def appearance(before, after, out_dir, dpi=144):
    """Coarse and fine metrics; reviewers inspect intentional detail changes."""
    docs = [pdfium.PdfDocument(str(p)) for p in (before, after)]
    if len(docs[0]) != len(docs[1]):
        raise AssertionError('Page count changed')
    results = []
    for i in range(len(docs[0])):
        images = []
        for doc in docs:
            page = doc[i]
            bitmap = page.render(scale=dpi / 72)
            images.append(bitmap.to_pil().convert('RGB'))
            bitmap.close(); page.close()
        a, b = images
        if a.size != b.size:
            raise AssertionError('Render dimensions changed')
        diff = np.abs(np.asarray(a).astype(float) - np.asarray(b).astype(float))
        coarse = np.abs(np.asarray(a.filter(ImageFilter.GaussianBlur(1.4))).astype(float)
                        - np.asarray(b.filter(ImageFilter.GaussianBlur(1.4))).astype(float))
        out_dir.mkdir(parents=True, exist_ok=True)
        a.save(out_dir / f'page{i+1}_native.png')
        b.save(out_dir / f'page{i+1}_restored.png')
        Image.fromarray(np.minimum(diff * 4, 255).astype('uint8')).save(
            out_dir / f'page{i+1}_difference_x4.png')
        results.append({'page': i + 1, 'dpi': dpi,
                        'mean_absolute_rgb_difference': float(diff.mean()),
                        'blurred_mean_absolute_rgb_difference': float(coarse.mean()),
                        'fraction_pixels_any_channel_diff_gt32':
                            float((diff.max(axis=2) > 32).mean()),
                        'blurred_fraction_pixels_any_channel_diff_gt32':
                            float((coarse.max(axis=2) > 32).mean())})
    for doc in docs:
        doc.close()
    return results


def audit(before, after, out_dir, asset_dir=None):
    old, new = inspect_pdf(before), inspect_pdf(after)
    if len(old) != len(new):
        raise AssertionError('Page count changed')
    checks, changed = [], []
    for i, (a, b) in enumerate(zip(old, new), 1):
        checks.append({'page': i,
                       **{key + '_unchanged': a[key] == b[key] for key in
                          ('mediabox', 'cropbox', 'rotation', 'fonts', 'text')},
                       # Restoration may wrap an image in a new clipped Form.
                       # Every pre-existing content stream must remain exact.
                       'original_content_streams_unchanged': all(
                           b['streams'].get(k) == v for k, v in a['streams'].items()),
                       'original_xobject_resource_names_preserved':
                           a['xobjects'].keys() <= b['xobjects'].keys()})
        changed.extend({'page': i, 'image': key, 'before': a['images'].get(key), 'after': item}
                       for key, item in b['images'].items() if a['images'].get(key) != item)
    result = {'before': str(before), 'after': str(after), 'invariants': checks,
              'all_nonimage_invariants_pass': all(all(v for k, v in c.items() if k != 'page')
                                                  for c in checks),
              'changed_images': changed,
              'appearance': appearance(before, after, Path(out_dir))}
    if asset_dir:
        result['exact_source_pixel_matches'] = source_matches(after, asset_dir)
    Path(out_dir, 'qa.json').write_text(json.dumps(result, indent=2), encoding='utf-8')
    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('before', type=Path)
    parser.add_argument('after', type=Path)
    parser.add_argument('--out', required=True, type=Path)
    parser.add_argument('--assets', type=Path)
    args = parser.parse_args()
    result = audit(args.before, args.after, args.out, args.assets)
    print(json.dumps({'all_nonimage_invariants_pass': result['all_nonimage_invariants_pass'],
                      'changed_images': len(result['changed_images']),
                      'appearance': result['appearance']}, indent=2))
    if not result['all_nonimage_invariants_pass']:
        raise SystemExit(1)
