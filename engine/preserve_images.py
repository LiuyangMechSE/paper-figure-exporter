"""Restore original PPT raster samples inside PowerPoint's vector PDF.

No output pixels are interpolated.  A matched native image is replaced by a
clipped Form containing the original source image, so fractional crops, rotation,
alpha and the original PDF drawing order remain intact.  Unsupported effects,
ambiguous matches and unexpected masks are deliberately left untouched and
reported.  This module does not create detail absent from the PPT.
"""
from __future__ import annotations

from collections import defaultdict
from io import BytesIO
import hashlib
import math
from pathlib import Path
import posixpath
from zipfile import ZipFile
import xml.etree.ElementTree as ET

from PIL import Image, ImageChops, ImageStat
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (ArrayObject, ContentStream, DecodedStreamObject, EncodedStreamObject,
                           DictionaryObject, FloatObject, NameObject,
                           NumberObject, BooleanObject)

from pdf_images import lossless_stream

NS = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
I = (1., 0., 0., 1., 0., 0.)
FLIP_Y = (1., 0., 0., -1., 0., 1.)


def _mul(a, b):
    return (a[0]*b[0]+a[2]*b[1], a[1]*b[0]+a[3]*b[1],
            a[0]*b[2]+a[2]*b[3], a[1]*b[2]+a[3]*b[3],
            a[0]*b[4]+a[2]*b[5]+a[4], a[1]*b[4]+a[3]*b[5]+a[5])


def _inv(a):
    d = a[0]*a[3]-a[1]*a[2]
    if abs(d) < 1e-18:
        raise ValueError('Singular image transformation')
    return (a[3]/d, -a[1]/d, -a[2]/d, a[0]/d,
            (a[2]*a[5]-a[3]*a[4])/d, (a[1]*a[4]-a[0]*a[5])/d)


def _point(m, x, y):
    return m[0]*x+m[2]*y+m[4], m[1]*x+m[3]*y+m[5]


def _bounds(m):
    p = [_point(m, x, y) for x, y in ((0, 0), (1, 0), (0, 1), (1, 1))]
    return [min(x for x, y in p), min(y for x, y in p),
            max(x for x, y in p), max(y for x, y in p)]


def _translate(x, y):
    return (1., 0., 0., 1., x, y)


def _scale(x, y):
    return (x, 0., 0., y, 0., 0.)


def _local(n):
    return n.tag.rsplit('}', 1)[-1]


def _num(n, key, default=0):
    return float(n.get(key, default)) if n is not None else default


def _shape_id(node):
    for child in node:
        if _local(child).startswith('nv'):
            for n in child:
                if _local(n) == 'cNvPr':
                    return int(n.get('id'))
    return None


def _shape_matrix(node, group=False):
    x = next((node.find(path, NS) for path in
              ('a:xfrm', 'p:xfrm', 'p:spPr/a:xfrm', 'p:grpSpPr/a:xfrm')
              if node.find(path, NS) is not None), None)
    if x is None:
        raise ValueError('Picture/group has no explicit transform')
    off, ext = x.find('a:off', NS), x.find('a:ext', NS)
    ox, oy = _num(off, 'x'), _num(off, 'y')
    ex, ey = _num(ext, 'cx'), _num(ext, 'cy')
    angle = math.radians(_num(x, 'rot')/60000)
    rot = (math.cos(angle), math.sin(angle), -math.sin(angle), math.cos(angle), 0, 0)
    fx = -1 if x.get('flipH') in ('1', 'true') else 1
    fy = -1 if x.get('flipV') in ('1', 'true') else 1
    m = _mul(_translate(ox+ex/2, oy+ey/2), rot)
    if group:
        off, ext = x.find('a:chOff', NS), x.find('a:chExt', NS)
        cx, cy = _num(off, 'x'), _num(off, 'y')
        cw, ch = _num(ext, 'cx', ex), _num(ext, 'cy', ey)
        if not cw or not ch:
            raise ValueError('Zero group coordinate extent')
        return _mul(_mul(m, _scale(fx*ex/cw, fy*ey/ch)), _translate(-cx-cw/2, -cy-ch/2))
    return _mul(_mul(m, _scale(fx*ex, fy*ey)), _translate(-.5, -.5))


def _rels(z, part):
    name = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part)+'.rels')
    if name not in z.namelist():
        return {}
    result = {}
    for r in ET.fromstring(z.read(name)):
        if r.get('TargetMode') != 'External':
            target = r.get('Target', '')
            result[r.get('Id')] = (target.lstrip('/') if target.startswith('/') else
                                  posixpath.normpath(posixpath.join(posixpath.dirname(part), target)))
    return result


def _source_pictures(source, slide_number, shape_ids, canvas_pt, page_height):
    pictures, warnings = [], []
    selected = set(map(int, shape_ids))
    with ZipFile(source) as z:
        presentation = ET.fromstring(z.read('ppt/presentation.xml'))
        ids = list(presentation.find('p:sldIdLst', NS))
        if not 1 <= slide_number <= len(ids):
            raise ValueError('Slide number is outside presentation')
        rels = _rels(z, 'ppt/presentation.xml')
        part = rels[ids[slide_number-1].get('{'+NS['r']+'}id')]
        rels = _rels(z, part)
        root = ET.fromstring(z.read(part))
        # PPT top-left EMUs -> native PDF bottom-left points. Canvas translation
        # was performed on the private PPT copy before PowerPoint rendered it.
        to_pdf = (1/12700., 0., 0., -1/12700., -canvas_pt[0], page_height+canvas_pt[1])

        def walk(node, parent=I, top_id=None, ancestors=()):
            tag = _local(node)
            if tag == 'AlternateContent':
                choices = [c for c in node if _local(c) == 'Choice']
                branch = choices[0] if choices else next((c for c in node if _local(c) == 'Fallback'), None)
                if branch is not None:
                    for c in branch:
                        walk(c, parent, top_id, ancestors)
                return
            if tag not in ('grpSp', 'sp', 'pic', 'graphicFrame', 'cxnSp'):
                return
            sid = _shape_id(node)
            top_id = sid if top_id is None else top_id
            if top_id not in selected:
                return
            if tag == 'grpSp':
                try:
                    group_matrix = _mul(parent, _shape_matrix(node, True))
                except ValueError as exc:
                    warnings.append(f'Shape {sid}: {exc}; group pictures cannot be restored.')
                    return
                for child in node:
                    walk(child, group_matrix, top_id, ancestors+(sid,))
                return
            fill = node.find('p:blipFill', NS)
            if fill is None:
                fill = node.find('p:spPr/a:blipFill', NS)
            if fill is None:
                return
            rec = {'shape_id': sid, 'top_level_shape_id': top_id, 'group_ids': list(ancestors),
                   'unsupported': [], 'shape_kind': tag}
            pictures.append(rec)
            try:
                if tag != 'pic':
                    rec['unsupported'].append('Raster shape fill (not a native picture)')
                blip = fill.find('a:blip', NS)
                media = rels.get(blip.get('{'+NS['r']+'}embed')) if blip is not None else None
                rec['media'] = media
                if not media:
                    raise ValueError('Linked/unresolved raster source')
                raw = z.read(media)
                im = Image.open(BytesIO(raw))
                rec['format'] = im.format
                rec['source_mode'] = im.mode
                if im.info.get('icc_profile'):
                    rec['unsupported'].append('Embedded ICC color profile; native color-managed rendering retained')
                if im.format not in ('PNG', 'JPEG', 'TIFF', 'BMP', 'GIF'):
                    raise ValueError('Unsupported source format '+str(im.format))
                if getattr(im, 'n_frames', 1) > 1:
                    rec['unsupported'].append('Animated/multi-frame source')
                if im.mode not in ('1', 'L', 'LA', 'P', 'RGB', 'RGBA'):
                    rec['unsupported'].append('Unsupported source color mode '+im.mode)
                if im.getexif().get(274, 1) != 1:
                    rec['unsupported'].append('EXIF image orientation')
                rec['image'] = im.convert('RGBA')
                rec['raw'] = raw
                rec['pixels'] = list(im.size)
                rec['sha256'] = hashlib.sha256(raw).hexdigest()
                effects = []
                for child in blip:
                    if _local(child) != 'extLst':
                        effects.append(_local(child))
                for n in node.iter():
                    if _local(n) in ('effectLst', 'effectDag') and len(n):
                        effects.extend(_local(e) for e in n)
                if effects:
                    rec['unsupported'].append('Picture effects: '+', '.join(sorted(set(effects))))
                geom = node.find('p:spPr/a:prstGeom', NS)
                if geom is not None and geom.get('prst') != 'rect':
                    rec['unsupported'].append('Non-rectangular picture geometry: '+str(geom.get('prst')))
                if node.find('p:spPr/a:custGeom', NS) is not None:
                    rec['unsupported'].append('Custom picture geometry')
                if fill.find('a:tile', NS) is not None:
                    rec['unsupported'].append('Tiled picture fill')
                dest = fill.find('a:stretch/a:fillRect', NS)
                if dest is not None and any(_num(dest, k) != 0 for k in ('l', 't', 'r', 'b')):
                    rec['unsupported'].append('Non-default fill rectangle')
                crop = fill.find('a:srcRect', NS)
                l, t, r, b = [_num(crop, k)/100000 for k in ('l', 't', 'r', 'b')]
                rec['crop_fractions'] = [l, t, r, b]
                vw, vh = 1-l-r, 1-t-b
                if vw <= 0 or vh <= 0:
                    raise ValueError('Invalid crop rectangle')
                frame = _mul(to_pdf, _mul(parent, _shape_matrix(node)))
                rec['frame'] = frame
                rec['matrix'] = _mul(frame, (1/vw, 0, 0, 1/vh, -l/vw, -t/vh))
            except Exception as exc:
                rec['unsupported'].append(str(exc))

        for node in root.find('p:cSld/p:spTree', NS):
            walk(node)
    return pictures, warnings


def _placements(reader):
    result, warnings = [], []
    for pi, page in enumerate(reader.pages):
        def walk(stream, resources, ctm=I, seen=()):
            resources = resources.get_object() if resources else DictionaryObject()
            stack = []
            for args, op in ContentStream(stream, reader).operations:
                if op == b'q':
                    stack.append(ctm)
                elif op == b'Q':
                    if stack:
                        ctm = stack.pop()
                elif op == b'cm':
                    ctm = _mul(ctm, tuple(float(x) for x in args))
                elif op == b'Do':
                    ref = resources.get('/XObject', {}).get(args[0])
                    if ref is None:
                        warnings.append('Missing PDF XObject '+str(args[0]))
                        continue
                    obj = ref.get_object()
                    key = getattr(ref, 'idnum', id(obj))
                    if obj.get('/Subtype') == '/Image' and not obj.get('/ImageMask', False):
                        result.append({'xref': key, 'name': str(args[0]), 'ctm': ctm,
                                       'ref': ref, 'object': obj, 'page': pi+1,
                                       'pixels': [int(obj['/Width']), int(obj['/Height'])]})
                    elif obj.get('/Subtype') == '/Form':
                        if key in seen:
                            warnings.append('Recursive PDF Form')
                            continue
                        walk(obj, obj.get('/Resources', resources),
                             _mul(ctm, tuple(float(x) for x in obj.get('/Matrix', I))), seen+(key,))
                elif op == b'INLINE IMAGE':
                    warnings.append('Inline PDF image retained; original pixel mapping unavailable.')
        walk(page.get('/Contents'), page.get('/Resources'))
    return result, warnings


def _ppi(pixels, matrix):
    x, y = math.hypot(matrix[0], matrix[1]), math.hypot(matrix[2], matrix[3])
    return [pixels[0]*72/x if x else None, pixels[1]*72/y if y else None]


def _appearance(source, native, domain, return_images=False):
    """Comparison ONLY; resampling here is never embedded in the output PDF."""
    size = native.size
    # Bound validation cost; tiny native objects are kept at their native size.
    if max(size) > 400:
        factor = 400/max(size)
        size = tuple(max(1, round(v*factor)) for v in size)
        native = native.resize(size, Image.Resampling.LANCZOS)
    sw, sh = source.size
    a, b, c, d, e, f = domain
    box = _bounds(domain)
    box = (box[0]*sw, box[1]*sh, box[2]*sw, box[3]*sh)
    integer_box = (math.floor(box[0]), math.floor(box[1]), math.ceil(box[2]), math.ceil(box[3]))
    region = source.crop(integer_box)
    fraction = (box[0]-integer_box[0], box[1]-integer_box[1],
                box[2]-integer_box[0], box[3]-integer_box[1])
    # resize(box=...) supplies a true antialias filter for strong downsampling;
    # affine bicubic alone aliases textured photographs and rejects valid matches.
    direct = abs(a)+abs(d) >= abs(b)+abs(c)
    predicted = region.resize(size if direct else size[::-1], Image.Resampling.LANCZOS, box=fraction)
    if not direct:
        predicted = predicted.transpose(Image.Transpose.TRANSPOSE)
    if (a if direct else b) < 0:
        predicted = predicted.transpose(Image.Transpose.FLIP_LEFT_RIGHT)
    if (d if direct else c) < 0:
        predicted = predicted.transpose(Image.Transpose.FLIP_TOP_BOTTOM)
    native = native.convert('RGBA')
    alpha_error = ImageStat.Stat(ImageChops.difference(predicted.getchannel('A'), native.getchannel('A'))).mean[0]/255
    bg = Image.new('RGBA', size, 'white')
    p = Image.alpha_composite(bg, predicted).convert('RGB')
    n = Image.alpha_composite(bg, native).convert('RGB')
    color_error = sum(ImageStat.Stat(ImageChops.difference(p, n)).mean)/(255*3)
    if return_images:
        return color_error, alpha_error, predicted, native
    return color_error, alpha_error


def _slide_shows_transparency(pic, best, placement, preview_path, slide_size_pt, canvas_pt, page_height):
    """Decide, from PowerPoint's own on-screen slide image, whether a picture's
    transparent areas are really transparent on the slide.

    PowerPoint's PDF export sometimes fills a picture's transparent areas
    (notably TIFF/PNG alpha, sometimes converting the picture to JPEG). The
    slide preview exported by PowerPoint shows what the slide really looks like.
    Returns a dict with 'verdict' in {'transparent', 'opaque', None}; None
    means the evidence is insufficient, and the caller keeps the native raster.
    """
    from PIL import ImageFilter
    result = {'verdict': None}
    if preview_path is None or slide_size_pt is None or not Path(preview_path).exists():
        result['reason'] = 'no slide preview'
        return result
    a, b, c, d, e, f = placement['ctm']
    if abs(b) + abs(c) > 1e-6 * max(abs(a), abs(d), 1e-6):
        result['reason'] = 'rotated placement'
        return result
    x0, y0, x1, y1 = _bounds(placement['ctm'])
    sw, sh = slide_size_pt
    left, right = x0 + canvas_pt[0], x1 + canvas_pt[0]
    top, bottom = page_height - y1 + canvas_pt[1], page_height - y0 + canvas_pt[1]
    if left < -1 or top < -1 or right > sw + 1 or bottom > sh + 1:
        result['reason'] = 'picture extends beyond the slide preview'
        return result
    preview = Image.open(preview_path).convert('RGB')
    kx, ky = preview.width / sw, preview.height / sh
    box = (max(0, left*kx), max(0, top*ky), min(preview.width, right*kx), min(preview.height, bottom*ky))
    size = (max(8, round(box[2]-box[0])), max(8, round(box[3]-box[1])))
    if size[0] * size[1] < 400:
        result['reason'] = 'picture too small in the slide preview'
        return result
    shown = preview.resize(size, Image.Resampling.LANCZOS, box=box)
    # Display orientation -> native image orientation (same flips as the PDF placement).
    if a < 0:
        shown = shown.transpose(Image.Transpose.FLIP_LEFT_RIGHT)
    if d < 0:
        shown = shown.transpose(Image.Transpose.FLIP_TOP_BOTTOM)
    _, _, predicted, native = _appearance(pic['image'], placement['native'], best['domain'], return_images=True)
    predicted, native = predicted.resize(size, Image.Resampling.LANCZOS), native.convert('RGB').resize(size, Image.Resampling.LANCZOS)
    alpha = predicted.getchannel('A')
    # Erode both masks so antialiased edges do not blur the evidence.
    clear = alpha.point(lambda v: 255 if v < 26 else 0).filter(ImageFilter.MinFilter(3))
    solid = alpha.point(lambda v: 255 if v > 229 else 0).filter(ImageFilter.MinFilter(3))
    area = size[0] * size[1]
    clear_fraction = ImageStat.Stat(clear).sum[0] / 255 / area
    solid_fraction = ImageStat.Stat(solid).sum[0] / 255 / area
    result.update(clear_fraction=round(clear_fraction, 4), solid_fraction=round(solid_fraction, 4))
    if clear_fraction < .05 or solid_fraction < .05:
        result['reason'] = 'not enough transparent and opaque area to compare'
        return result

    def masked_diff(x, y, mask):
        diff = ImageChops.difference(x, y).convert('L')
        return ImageStat.Stat(diff, mask).mean[0] / 255

    source_rgb = Image.alpha_composite(Image.new('RGBA', size, 'white'), predicted).convert('RGB')
    content_match = masked_diff(source_rgb, native, solid)   # source vs PDF where the source is opaque
    mapping_check = masked_diff(shown, native, solid)        # slide vs PDF where the source is opaque
    clear_vs_native = masked_diff(shown, native, clear)      # slide vs PDF where the source is transparent
    result.update(content_match=round(content_match, 4), mapping_check=round(mapping_check, 4),
                  clear_vs_native=round(clear_vs_native, 4))
    if content_match > .055 or mapping_check > .12:
        result['reason'] = 'opaque areas do not agree well enough to judge'
        return result
    if clear_vs_native > max(.10, 2.5 * mapping_check):
        result['verdict'] = 'transparent'
    elif clear_vs_native < .06:
        result['verdict'] = 'opaque'
    else:
        result['reason'] = 'inconclusive'
    return result


def _unmatte(native, obj):
    """pypdf's PIL image includes alpha but does not undo PDF /Matte colors."""
    mask = obj.get('/SMask')
    matte = mask.get_object().get('/Matte') if mask is not None else None
    if matte is None:
        return native
    if len(matte) != 3 or obj.get('/ColorSpace') != '/DeviceRGB':
        return native
    colors = [float(v)*255 for v in matte]
    raw = bytearray(native.convert('RGBA').tobytes())
    for offset in range(0, len(raw), 4):
        alpha = raw[offset+3]/255
        if alpha:
            for channel in range(3):
                value = (raw[offset+channel]-(1-alpha)*colors[channel])/alpha
                raw[offset+channel] = max(0, min(255, round(value)))
    return Image.frombytes('RGBA', native.size, bytes(raw))


def _candidate(pic, placement, native):
    if 'matrix' not in pic:
        return None
    # Native image's top-left pixel square expressed in source top-left units.
    domain = _mul(_inv(pic['matrix']), _mul(placement['ctm'], FLIP_Y))
    a, b, c, d, e, f = domain
    # Do not attempt to reverse rasterized arbitrary rotation/shear/effects.
    axis = min(abs(b)+abs(c), abs(a)+abs(d))
    if axis > .003*max(abs(a), abs(b), abs(c), abs(d), .001):
        return None
    box = _bounds(domain)
    l, t, r, bottom = pic['crop_fractions']
    expected = [max(0, l), max(0, t), min(1, 1-r), min(1, 1-bottom)]
    tolx = max(.002, 1.5/pic['pixels'][0])
    toly = max(.002, 1.5/pic['pixels'][1])
    # PDF rounding/downsampling can extend one output pixel beyond the crop.
    tolx += 1.1*abs(box[2]-box[0])/placement['pixels'][0]
    toly += 1.1*abs(box[3]-box[1])/placement['pixels'][1]
    err = max(abs(box[k]-expected[k])/(tolx if k % 2 == 0 else toly) for k in range(4))
    if err > 2.5:
        return None
    color_error, alpha_error = _appearance(pic['image'], native, domain)
    return {'picture': pic, 'domain': domain, 'geometry_error_units': err,
            'color_error': color_error, 'alpha_error': alpha_error,
            'score': color_error+alpha_error+err*.002}


def _image_stream(writer, pic):
    im = pic['image']
    # Original JPEG codestreams are embedded byte-for-byte: no decode/re-encode
    # loss and no inflation of large photograph panels into raw RGB Flate data.
    source_jpeg = pic['format'] == 'JPEG' and pic['source_mode'] == 'RGB'
    if source_jpeg:
        stream = EncodedStreamObject()
        stream._data = pic['raw']
        stream[NameObject('/Filter')] = NameObject('/DCTDecode')
        stream.update({NameObject('/Type'): NameObject('/XObject'),
                       NameObject('/Subtype'): NameObject('/Image'),
                       NameObject('/Width'): NumberObject(im.width),
                       NameObject('/Height'): NumberObject(im.height),
                       NameObject('/BitsPerComponent'): NumberObject(8),
                       NameObject('/ColorSpace'): NameObject('/DeviceRGB'),
                       NameObject('/Interpolate'): BooleanObject(False)})
    else:
        # Exact samples, PNG-style prediction + deflate (lossless, compact).
        stream = lossless_stream(im.convert('RGB'))
    alpha = im.getchannel('A')
    if alpha.getextrema() != (255, 255):
        stream[NameObject('/SMask')] = writer._add_object(lossless_stream(alpha))
    return writer._add_object(stream)


def _replace_with_form(writer, placement, pic, cache):
    key = pic['sha256']
    if key not in cache:
        cache[key] = _image_stream(writer, pic)
    # Source top-left -> original PDF image unit square. Full source pixels are
    # placed without resizing; the original native crop is the Form's BBox.
    local_matrix = _mul(_inv(placement['ctm']), _mul(pic['matrix'], FLIP_Y))
    form = DecodedStreamObject()
    # PDF numeric tokens do NOT permit Python's scientific notation.
    matrix_text = ' '.join(format(v, '.12f').rstrip('0').rstrip('.') or '0' for v in local_matrix)
    form.set_data(('q 0 0 1 1 re W n '+matrix_text+' cm /OriginalPixels Do Q\n').encode('ascii'))
    form.update({NameObject('/Type'): NameObject('/XObject'),
                 NameObject('/Subtype'): NameObject('/Form'),
                 NameObject('/FormType'): NumberObject(1),
                 NameObject('/BBox'): ArrayObject([NumberObject(0), NumberObject(0), NumberObject(1), NumberObject(1)]),
                 NameObject('/Resources'): DictionaryObject({NameObject('/XObject'): DictionaryObject({NameObject('/OriginalPixels'): cache[key]})})})
    obj = placement['ref'].get_object()
    # The writer owns these cloned objects. Replacing the resource object keeps
    # every native content stream, its outer CTM, clipping and paint order intact.
    writer._objects[obj.indirect_reference.idnum-1] = form
    form.indirect_reference = obj.indirect_reference
    return local_matrix


def restore_images(source_pptx, input_pdf, output_pdf, slide_number, shape_ids, canvas_pt,
                   slide_preview=None, slide_size_pt=None):
    """Write a new PDF, returning JSON-serializable evidence and limitations.

    ``canvas_pt`` is [left,top,right,bottom] in the original slide's points.
    Only the selected top-level shapes (including their group descendants) are
    eligible. Existing output paths are refused. Every placed native image and
    every unmatched source raster is accounted for in the returned report.
    """
    source_pptx, input_pdf, output_pdf = map(Path, (source_pptx, input_pdf, output_pdf))
    if output_pdf.exists():
        raise FileExistsError('Refusing to overwrite '+str(output_pdf))
    reader = PdfReader(input_pdf)
    if len(reader.pages) != 1:
        raise ValueError('Expected one native PDF page per figure')
    page = reader.pages[0]
    if page.get('/Rotate', 0) or float(page.get('/UserUnit', 1)) != 1:
        raise ValueError('Rotated pages or non-default PDF UserUnit are unsupported')
    pictures, warnings = _source_pictures(source_pptx, int(slide_number), shape_ids,
                                          list(canvas_pt), float(page.mediabox.height))
    writer = PdfWriter(clone_from=reader)
    placements, pdf_warnings = _placements(writer)
    warnings.extend(pdf_warnings)
    decoded = {}
    # Public pypdf API applies native soft masks; keys may represent nested Forms.
    for image_key in reader.pages[0].images.keys():
        try:
            item = reader.pages[0].images[image_key]
            if item.indirect_reference:
                decoded[item.indirect_reference.idnum] = item.image.copy().convert('RGBA')
        except Exception as exc:
            warnings.append('Native PDF image '+str(image_key)+' could not be decoded for validation: '+str(exc))
    # Writer cloning changes object numbers; map original placements by order.
    original_placements, _ = _placements(reader)
    grouped = defaultdict(list)
    for placement, original in zip(placements, original_placements):
        native = decoded.get(original['xref'])
        placement['native'] = _unmatte(native, original['object']) if native is not None else None
        grouped[placement['xref']].append(placement)
    report = {'slide': int(slide_number), 'method': 'original_pixels_clipped_form',
              'resampling': False, 'source_picture_count': len(pictures),
              'native_image_placements': len(placements), 'restored_resources': 0,
              'retained_resources': 0, 'images': [], 'warnings': warnings}
    matched, cache = set(), {}
    for xref, uses in grouped.items():
        placement = uses[0]
        rec = {'pdf_xref': xref, 'pdf_resource': placement['name'], 'placements': len(uses),
               'native_pixels': placement['pixels'], 'before_ppi': _ppi(placement['pixels'], placement['ctm']),
               'status': 'retained_native', 'warnings': []}
        report['images'].append(rec)
        candidates = []
        if placement['native'] is not None:
            for pic in pictures:
                if 'image' in pic:
                    candidate = _candidate(pic, placement, placement['native'])
                    if candidate is not None:
                        candidates.append(candidate)
        candidates.sort(key=lambda c: c['score'])
        if not candidates:
            rec['warnings'].append('No source picture passed geometry matching; native raster retained.')
        else:
            best = candidates[0]
            pic = best['picture']
            rec.update({'shape_id': pic['shape_id'], 'source_media': pic['media'], 'source_pixels': pic['pixels'],
                        'appearance_mean_error': best['color_error'], 'alpha_mean_error': best['alpha_error'],
                        'geometry_error_units': best['geometry_error_units'],
                        'source_ppi': _ppi(pic['pixels'], pic['matrix'])})
            matched.add(pic['shape_id'])
            ambiguous = len(candidates) > 1 and candidates[1]['score']-best['score'] < .015
            if ambiguous:
                rec['warnings'].append('Ambiguous source picture match; native raster retained.')
            rec['warnings'].extend(pic['unsupported'])
            if best['color_error'] > .055 or best['alpha_error'] > .055:
                check = _slide_shows_transparency(pic, best, placement, slide_preview, slide_size_pt,
                                                  canvas_pt, float(page.mediabox.height))
                rec['slide_transparency_check'] = check
                if check['verdict'] == 'transparent' and not ambiguous:
                    rec['notes'] = ['PowerPoint\'s PDF export filled this picture\'s transparent areas; the slide shows them '
                                    'transparent. Original pixels and transparency restored to match the slide.']
                else:
                    rec['warnings'].append('Native raster appearance/alpha differs from source (effects, mask or mapping); retained.')
            if '/Mask' in placement['object']:
                rec['warnings'].append('Native color-key/stencil mask is unsupported; retained.')
            shared_rounding_error = 0.
            for other in uses[1:]:
                other_candidates = []
                if other['native'] is not None:
                    for other_pic in pictures:
                        if 'image' in other_pic:
                            cand = _candidate(other_pic, other, other['native'])
                            if cand is not None:
                                other_candidates.append(cand)
                other_candidates.sort(key=lambda c: c['score'])
                cand = other_candidates[0] if other_candidates else None
                mismatch = float('inf')
                if cand is not None:
                    first_local = _mul(_inv(placement['ctm']), _mul(pic['matrix'], FLIP_Y))
                    other_local = _mul(_inv(other['ctm']), _mul(cand['picture']['matrix'], FLIP_Y))
                    first_global = _mul(other['ctm'], first_local)
                    other_global = _mul(other['ctm'], other_local)
                    mismatch = max(math.dist(_point(first_global,x,y), _point(other_global,x,y))
                                   for x,y in ((0,0),(1,0),(0,1),(1,1)))
                if (cand is None or cand['picture']['sha256'] != pic['sha256'] or
                    cand['picture']['unsupported'] or cand['color_error'] > .055 or cand['alpha_error'] > .055 or
                    (len(other_candidates) > 1 and other_candidates[1]['score']-cand['score'] < .015) or
                    mismatch > .01):
                    rec['warnings'].append('Shared image resource has inconsistent placement/crops; native raster retained.')
                    break
                shared_rounding_error = max(shared_rounding_error, mismatch)
                matched.add(cand['picture']['shape_id'])
            if len(uses) > 1:
                rec['shared_transform_rounding_error_pt'] = shared_rounding_error
            if not rec['warnings']:
                local = _replace_with_form(writer, placement, pic, cache)
                rec.update({'status': 'restored_original_pixels', 'retained_pixels': pic['pixels'],
                            'after_ppi': _ppi(pic['pixels'], pic['matrix']),
                            'local_image_matrix': list(local), 'source_sha256': pic['sha256'],
                            'fractional_crop': 'Exact affine placement of original pixels within native unit-square clip'})
                report['restored_resources'] += 1
        if rec['status'] == 'retained_native':
            rec['after_ppi'] = rec['before_ppi']
            report['retained_resources'] += 1
    report['unmatched_source_pictures'] = [
        {'shape_id': p['shape_id'], 'media': p.get('media'), 'warnings': p['unsupported'] or
         ['No native PDF image matched this source picture; it may be absent, vector-rendered or combined.']}
        for p in pictures if p['shape_id'] not in matched]
    for p in report['unmatched_source_pictures']:
        warnings.append('Unmatched source picture '+str(p['shape_id'])+': '+'; '.join(p['warnings']))
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    with output_pdf.open('xb') as handle:
        writer.write(handle)
    return report
