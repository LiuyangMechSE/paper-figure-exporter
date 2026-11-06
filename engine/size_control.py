"""Bring a figure PDF under a file-size limit without lossy compression.

Ladder, stopping as soon as the file fits:
  1. Lossless: re-encode pictures with PNG-style prediction (identical pixels),
     merge duplicate objects, compress page content at the maximum level.
  2. Reduce only pictures whose resolution at their final printed size is above
     a cap, lowering the cap step by step but NEVER below ``min_ppi``.
     Reduced pictures are resampled once from the original pixels (Lanczos,
     premultiplied alpha) and stored losslessly. No JPEG is ever written;
     original JPEG pictures are kept byte for byte unless a reduced lossless
     version is genuinely smaller.
Text, vector drawings, fonts, page size and layout are never changed. The
untouched input is kept by the caller (``keep_copy``) whenever anything changes.
"""
from __future__ import annotations

from io import BytesIO
import hashlib
import math
import os
from pathlib import Path
import shutil

from PIL import Image
from pypdf import PdfReader, PdfWriter
from pypdf.generic import IndirectObject, NameObject

from image_quality import assess_pdf
from pdf_images import lossless_stream

MB = 1024 * 1024
# Standard filters whose decoding is exact; any chain of them can be re-encoded losslessly.
LOSSLESS_FILTERS = {'/FlateDecode', '/ASCII85Decode', '/ASCIIHexDecode', '/LZWDecode', '/RunLengthDecode'}


def _key(obj):
    raw = obj._data if hasattr(obj, '_data') else b''
    return (int(obj.get('/Width', 0)), int(obj.get('/Height', 0)), hashlib.sha1(raw).hexdigest())


def _image_refs(writer):
    """Indirect references of every image XObject reachable from the pages."""
    found, seen = {}, set()

    def visit(resources):
        resources = resources.get_object() if resources is not None else None
        if not resources or '/XObject' not in resources:
            return
        for ref in resources['/XObject'].get_object().values():
            if not isinstance(ref, IndirectObject) or ref.idnum in seen:
                continue
            seen.add(ref.idnum)
            obj = ref.get_object()
            if obj.get('/Subtype') == '/Image':
                found[ref.idnum] = ref
            elif obj.get('/Subtype') == '/Form':
                visit(obj.get('/Resources'))

    for page in writer.pages:
        visit(page.get('/Resources'))
    return found


def _decode(obj):
    """Return (PIL image in RGB/L, alpha or None) or None when not safely editable."""
    if obj.get('/ImageMask') or obj.get('/Mask') is not None or obj.get('/Decode') is not None:
        return None
    if int(obj.get('/BitsPerComponent', 0)) != 8:
        return None
    space = obj.get('/ColorSpace')
    space = space.get_object() if hasattr(space, 'get_object') else space
    if space == '/DeviceRGB':
        mode = 'RGB'
    elif space == '/DeviceGray':
        mode = 'L'
    elif isinstance(space, list) and len(space) == 2 and space[0] == '/ICCBased' and int(space[1].get_object().get('/N', 0)) in (1, 3):
        mode = 'RGB' if int(space[1].get_object()['/N']) == 3 else 'L'
    else:
        return None
    filters = obj.get('/Filter')
    filters = [filters] if isinstance(filters, str) or filters is None else list(filters)
    w, h = int(obj['/Width']), int(obj['/Height'])
    try:
        if filters == [None] or set(filters) <= LOSSLESS_FILTERS:
            im = Image.frombytes(mode, (w, h), obj.get_data())
        elif filters == ['/DCTDecode']:
            im = Image.open(BytesIO(obj._data))
            if im.mode != mode:
                return None  # e.g. CMYK JPEG: leave untouched
            im.load()
        else:
            return None
    except Exception:
        return None
    alpha = None
    if obj.get('/SMask') is not None:
        mask = obj['/SMask'].get_object()
        if (mask.get('/Matte') is not None or int(mask.get('/BitsPerComponent', 0)) != 8
                or mask.get('/ColorSpace') != '/DeviceGray' or mask.get('/Decode') is not None):
            return None
        try:
            alpha = Image.frombytes('L', (int(mask['/Width']), int(mask['/Height'])), mask.get_data())
        except Exception:
            return None
    return im, alpha


def _stored_bytes(obj):
    size = len(obj._data)
    if obj.get('/SMask') is not None:
        size += len(obj['/SMask'].get_object()._data)
    return size


def _resample(im, alpha, size):
    if alpha is None:
        return im.resize(size, Image.Resampling.LANCZOS), None
    if alpha.size != im.size:
        alpha = alpha.resize(im.size, Image.Resampling.LANCZOS)
    rgba = im.convert('RGB').copy()
    rgba.putalpha(alpha)
    # Premultiplied resampling avoids colour fringes at transparent edges.
    out = rgba.convert('RGBa').resize(size, Image.Resampling.LANCZOS).convert('RGBA')
    color = out.convert('RGB') if im.mode == 'RGB' else out.convert('RGB').convert('L')
    return color, out.getchannel('A')


def _build(reader, plan, ppi_cap, min_ppi, cache):
    """Clone the PDF and apply the lossless/reduction plan for one cap value."""
    writer = PdfWriter(clone_from=reader)
    changes = []
    for idnum, ref in _image_refs(writer).items():
        obj = ref.get_object()
        info = plan.get(_key(obj))
        if info is None:
            continue
        ppi = info['ppi']
        target = None
        if ppi_cap is not None and ppi is not None and ppi > ppi_cap * 1.0001:
            target = max(ppi_cap, min_ppi)
        w, h = int(obj['/Width']), int(obj['/Height'])
        if target is not None:
            factor = target / ppi
            size = (max(1, math.ceil(w * factor)), max(1, math.ceil(h * factor)))
        else:
            if info['already_compact']:
                continue
            size = (w, h)
        cache_key = (info['key'], size)
        if cache_key not in cache:
            im, alpha = info['decoded']
            if size != (w, h):
                im, alpha = _resample(im, alpha, size)
            elif alpha is not None and alpha.size != im.size:
                alpha = None  # keep the original mask object
            stream = lossless_stream(im)
            mask_stream = lossless_stream(alpha) if alpha is not None else None
            new_bytes = len(stream._data) + (len(mask_stream._data) if mask_stream is not None else 0)
            cache[cache_key] = (stream, mask_stream, new_bytes)
        stream, mask_stream, new_bytes = cache[cache_key]
        if new_bytes >= info['stored_bytes']:
            continue  # never make a picture bigger
        new = type(stream)()
        new._data = stream._data
        for k, v in stream.items():
            new[NameObject(k)] = v
        if obj.get('/ColorSpace') is not None and obj['/ColorSpace'] not in ('/DeviceRGB', '/DeviceGray'):
            new[NameObject('/ColorSpace')] = obj.raw_get('/ColorSpace')  # keep ICC profile
        if mask_stream is not None:
            m = type(mask_stream)()
            m._data = mask_stream._data
            for k, v in mask_stream.items():
                m[NameObject(k)] = v
            new[NameObject('/SMask')] = writer._add_object(m)
        elif obj.get('/SMask') is not None:
            new[NameObject('/SMask')] = obj.raw_get('/SMask')
        writer._objects[ref.idnum - 1] = new
        new.indirect_reference = ref
        changes.append(dict(pixels_before=[w, h], pixels_after=list(size),
                            ppi_before=None if ppi is None else round(ppi, 1),
                            ppi_after=None if ppi is None else round(ppi * size[0] / w, 1),
                            bytes_before=info['stored_bytes'], bytes_after=new_bytes,
                            resampled=size != (w, h)))
    for page in writer.pages:
        page.compress_content_streams(level=9)
    writer.compress_identical_objects()
    buffer = BytesIO()
    writer.write(buffer)
    return buffer.getvalue(), changes


def fit_file_size(pdf_path, max_bytes, min_ppi, keep_copy=None):
    """Make ``pdf_path`` at most ``max_bytes`` if possible; returns a JSON report.

    ``keep_copy``: where to copy the untouched input if anything changes.
    The output is the smallest version that respects ``min_ppi``, even when it
    still exceeds the limit (the report then says so).
    """
    path = Path(pdf_path)
    original = path.stat().st_size
    report = dict(limit_bytes=int(max_bytes), original_bytes=original, final_bytes=original,
                  meets_limit=original <= max_bytes, changed=False, method='unchanged',
                  ppi_cap=None, min_ppi=min_ppi, pictures=[], lossless_copy=None, notes=[])
    if original <= max_bytes:
        return report
    reader = PdfReader(str(path))
    quality = assess_pdf(path, target_ppi=min_ppi)
    ppi_by_object = {}
    for rec in quality['images']:
        if rec.get('inline') or rec.get('xref') is None or not rec.get('potentially_visible', True):
            continue
        ppi_by_object[rec['xref']] = min(ppi_by_object.get(rec['xref'], math.inf), rec['min_ppi'])
    plan = {}
    for idnum, ref in _image_refs(reader).items():
        obj = ref.get_object()
        decoded = _decode(obj)
        if decoded is None:
            continue
        parms = obj.get('/DecodeParms')
        parms = parms.get_object() if hasattr(parms, 'get_object') else parms
        if isinstance(parms, list):  # one DecodeParms entry per filter in a chain
            parms = next((x.get_object() for x in parms if hasattr(x, 'get_object') and x.get_object()), None)
        compact = obj.get('/Filter') == '/DCTDecode' or (hasattr(parms, 'get') and int(parms.get('/Predictor', 1)) >= 10)
        ppi = ppi_by_object.get(idnum)
        previous = plan.get(_key(obj))
        if previous is not None and previous['ppi'] is not None and ppi not in (None, math.inf):
            ppi = min(ppi, previous['ppi'])  # same pixels placed twice: respect the larger placement
        plan[_key(obj)] = dict(key=_key(obj), decoded=decoded, ppi=None if ppi in (None, math.inf) else ppi,
                               stored_bytes=_stored_bytes(obj), already_compact=compact)
    cache = {}
    attempts = []

    def attempt(cap):
        data, changes = _build(reader, plan, cap, min_ppi, cache)
        attempts.append(dict(ppi_cap=cap, bytes=len(data)))
        return data, changes

    best_data, best_changes = attempt(None)  # lossless only
    best_cap = None
    reducible = sorted({p['ppi'] for p in plan.values() if p['ppi'] and p['ppi'] > min_ppi * 1.0001}, reverse=True)
    if len(best_data) > max_bytes and reducible:
        # Feasibility first: the smallest file allowed at the PPI floor.
        floor_data, floor_changes = attempt(min_ppi)
        if len(floor_data) <= max_bytes:
            # Highest PPI cap that still fits, by bisection on real file sizes:
            # lo fits (the floor), hi does not (no reduction).
            best_data, best_changes, best_cap = floor_data, floor_changes, min_ppi
            lo, hi = min_ppi, reducible[0]
            for _ in range(6):
                if hi / lo < 1.04:
                    break
                mid = math.sqrt(lo * hi)
                data, changes = attempt(mid)
                if len(data) <= max_bytes:
                    best_data, best_changes, best_cap, lo = data, changes, mid, mid
                else:
                    hi = mid
        else:
            best_data, best_changes, best_cap = floor_data, floor_changes, min_ppi
            report['notes'].append(
                f'Even with every picture reduced to {min_ppi:g} PPI the file is {len(floor_data)/MB:.1f} MB. '
                'Options: allow a lower min_ppi, use a smaller figure width, or provide smaller source pictures.')
    elif len(best_data) > max_bytes:
        report['notes'].append('No picture is above the PPI floor, so it cannot be reduced further without going below it. '
                               'The remaining size comes from pictures at or below the floor, fonts or vector drawings.')
    if len(best_data) >= original:
        report['notes'].append('Lossless optimisation did not make this file smaller; the original was kept.')
        report['attempts'] = attempts
        return report
    if keep_copy is not None:
        keep_copy = Path(keep_copy)
        keep_copy.parent.mkdir(parents=True, exist_ok=True)
        if not keep_copy.exists():
            shutil.copy2(path, keep_copy)
        report['lossless_copy'] = f'{keep_copy.parent.name}/{keep_copy.name}'
    temp = path.with_name(path.name + '.partial')
    temp.write_bytes(best_data)
    os.replace(temp, path)
    report.update(final_bytes=len(best_data), meets_limit=len(best_data) <= max_bytes, changed=True,
                  method='lossless' if best_cap is None else 'lossless_plus_reduced_to_ppi_cap',
                  ppi_cap=None if best_cap is None else round(best_cap, 1), pictures=best_changes, attempts=attempts)
    return report
