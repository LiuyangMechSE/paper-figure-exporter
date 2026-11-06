"""Measure effective raster resolution from actual PDF image placements.

No PDF is edited. A vector PDF has no single DPI. Images are measured using the
full transformation matrix, including nested Forms, rotation and /UserUnit.
Masks are reported with their parent image, never as separate placed pictures.
Visibility uses conservative clipping bounds; it does not infer detail from a
nominal DPI tag, interpolate pixels, or resolve arbitrary artwork occlusion.
"""
from __future__ import annotations

import argparse
import json
import math
from pathlib import Path

from pypdf import PdfReader
from pypdf.generic import ContentStream

IDENTITY = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)


def _obj(value):
    return value.get_object() if hasattr(value, 'get_object') else value


def _bool(value):
    # pypdf BooleanObject(False) is truthy as a Python object.
    return bool(getattr(_obj(value), 'value', value))


def _matrix(outer, inner):
    """Column-vector affine composition outer @ inner, in PDF six-tuple form."""
    a, b = outer, inner
    return (a[0]*b[0]+a[2]*b[1], a[1]*b[0]+a[3]*b[1],
            a[0]*b[2]+a[2]*b[3], a[1]*b[2]+a[3]*b[3],
            a[0]*b[4]+a[2]*b[5]+a[4], a[1]*b[4]+a[3]*b[5]+a[5])


def _point(m, x, y):
    return (m[0]*x+m[2]*y+m[4], m[1]*x+m[3]*y+m[5])


def _bounds(m, box):
    p = [_point(m, x, y) for x in (box[0], box[2]) for y in (box[1], box[3])]
    return [min(x for x, y in p), min(y for x, y in p),
            max(x for x, y in p), max(y for x, y in p)]


def _intersect(a, b):
    return [max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3])]


def _area(box):
    return max(0.0, box[2]-box[0])*max(0.0, box[3]-box[1])


def _density(width, height, ctm, userunit):
    xpt = math.hypot(ctm[0], ctm[1])*userunit
    ypt = math.hypot(ctm[2], ctm[3])*userunit
    if min(width, height, xpt, ypt) <= 0:
        return None
    # Pixel basis vectors in physical PDF points. The largest singular value
    # gives the worst-direction physical pixel spacing, including shear.
    a, b = ctm[0]*userunit/width, ctm[1]*userunit/width
    c, d = ctm[2]*userunit/height, ctm[3]*userunit/height
    trace = a*a+b*b+c*c+d*d
    determinant = (a*d-b*c)**2
    largest = math.sqrt(max(0.0, (trace+math.sqrt(max(0.0, trace*trace-4*determinant)))/2))
    if largest <= 0 or abs(ctm[0]*ctm[3]-ctm[1]*ctm[2]) < 1e-16:
        return None
    return dict(dpi_x=width*72.0/xpt, dpi_y=height*72.0/ypt,
                directional_min_ppi=72.0/largest, width_pt=xpt, height_pt=ypt)


def _font_info(reference, default_name):
    f = _obj(reference)
    subtype = str(f.get('/Subtype', ''))
    embedded = subtype == '/Type3'
    descriptors = []
    if f.get('/FontDescriptor'):
        descriptors.append(_obj(f['/FontDescriptor']))
    for descendant in _obj(f.get('/DescendantFonts', [])):
        desc = _obj(descendant)
        if desc.get('/FontDescriptor'):
            descriptors.append(_obj(desc['/FontDescriptor']))
    embedded = embedded or any(any(k in d for k in ('/FontFile', '/FontFile2', '/FontFile3')) for d in descriptors)
    return dict(name=str(f.get('/BaseFont', default_name)), subtype=subtype, embedded=bool(embedded))


def assess_pdf(pdf_path, target_ppi=600):
    """Return serializable measurements and honest remedies at final PDF size.

    ``min_ppi`` is the limiting physical resolution of potentially visible
    pictures, including an attached image mask if it has fewer samples. A
    conservative significant-image subset is also exposed as ``min_large_ppi``;
    small icons are still present in ``images`` and ``below_target_images``.
    Required source dimensions mean genuinely higher-resolution source data,
    not interpolated copies. ``max_scale_at_target`` scales the CURRENT image
    placement (or whole figure) to the target without changing its pixels.
    """
    target_ppi = float(target_ppi)
    if not math.isfinite(target_ppi) or target_ppi <= 0:
        raise ValueError('target_ppi must be a finite positive number')
    source = Path(pdf_path)
    reader = PdfReader(str(source))
    images, pages, issues, fonts = [], [], [], {}
    counters = dict(path_ops=0, path_paint_ops=0, text_draw_ops=0, stencil_mask_draws=0)
    complete = True

    for page_number, page in enumerate(reader.pages, 1):
        page_images, page_issues, page_fonts = [], [], {}
        page_counts = dict(path_ops=0, path_paint_ops=0, text_draw_ops=0, stencil_mask_draws=0)
        userunit = float(page.get('/UserUnit', 1))
        if not math.isfinite(userunit) or userunit <= 0:
            raise ValueError(f'Page {page_number} has an invalid /UserUnit')
        mediabox = [float(x) for x in page.mediabox]
        cropbox = [float(x) for x in page.cropbox]
        initial_clip = _intersect(mediabox, cropbox)

        def issue(message, incomplete=False):
            nonlocal complete
            page_issues.append(message)
            if incomplete:
                complete = False

        for annotation in _obj(page.get('/Annots', [])):
            if _obj(annotation).get('/AP'):
                issue('Annotation appearance programs are not audited for raster images', True)
                break

        def add_image(obj, ctm, clip, clip_exact, alpha, gs_mask, where, name, xref, inline=False):
            nonlocal complete
            is_stencil = _bool(obj.get('/ImageMask', obj.get('/IM', False)))
            if is_stencil:
                page_counts['stencil_mask_draws'] += 1
                # A stencil painted with Do is visible sampled artwork. Only
                # masks attached to another image are folded into its record.
            try:
                w, h = int(obj.get('/Width', obj.get('/W'))), int(obj.get('/Height', obj.get('/H')))
            except (TypeError, ValueError):
                issue(f'{where}/{name}: image dimensions missing or invalid', True)
                return
            density = _density(w, h, ctm, userunit)
            if density is None:
                issue(f'{where}/{name}: degenerate image transform or dimensions', True)
                return
            bbox = _bounds(ctm, (0, 0, 1, 1))
            visible = _intersect(clip, bbox)
            visible_area = _area(visible)*userunit*userunit if alpha > 0 else 0.0
            spatial_masks, mask_unknown = [], False
            for mask_key in ('/SMask', '/Mask'):
                ref = obj.get(mask_key)
                if ref is None:
                    continue
                mask = _obj(ref)
                if isinstance(mask, (list, tuple)) or not hasattr(mask, 'get'):
                    continue  # Color-key masks have no independent sample grid.
                if mask.get('/Width') is not None and mask.get('/Height') is not None:
                    mw, mh = int(mask['/Width']), int(mask['/Height'])
                    md = _density(mw, mh, ctm, userunit)
                    if md:
                        spatial_masks.append(dict(kind=mask_key, width_px=mw, height_px=mh,
                                                  dpi_x=md['dpi_x'], dpi_y=md['dpi_y'],
                                                  min_ppi=md['directional_min_ppi'],
                                                  required_width_px=max(1, math.ceil(md['width_pt']/72*target_ppi-1e-9)),
                                                  required_height_px=max(1, math.ceil(md['height_pt']/72*target_ppi-1e-9)),
                                                  meets_target=md['directional_min_ppi']+1e-6 >= target_ppi))
                    else:
                        mask_unknown = True
                else:
                    mask_unknown = True
            limiting = min([density['directional_min_ppi']] + [m['min_ppi'] for m in spatial_masks])
            width_in, height_in = density['width_pt']/72.0, density['height_pt']/72.0
            small = min(w, h) < 32 or visible_area < 36.0
            is_visible = visible_area > 1e-10
            known_masks = not (mask_unknown or gs_mask)
            # Does not conflate higher encoded pixel counts with recoverable
            # image information. The restoration module records provenance.
            record = dict(page=page_number, image_index=len(images)+len(page_images)+1,
                          xref=xref, name=name, location=where, inline=inline,
                          stencil=is_stencil,
                          image_width_px=w, image_height_px=h, pixels=[w, h],
                          **density, min_ppi=limiting,
                          width_mm=width_in*25.4, height_mm=height_in*25.4,
                          ctm=list(ctm), userunit=userunit,
                          bbox_pt=[v*userunit for v in bbox],
                          visible_bbox_pt=[v*userunit for v in visible],
                          visible_area_pt2_upper_bound=visible_area,
                          potentially_visible=is_visible, clip_bounds_exact=clip_exact,
                          nonstroking_alpha=alpha, graphics_state_soft_mask=gs_mask,
                          small_element=small,
                          classification='small_icon_or_detail' if small else 'main_image',
                          filter=str(obj.get('/Filter', obj.get('/F', ''))),
                          bits_per_component=int(obj.get('/BitsPerComponent', obj.get('/BPC', 1))),
                          colorspace=str(obj.get('/ColorSpace', obj.get('/CS', ''))),
                          interpolate=_bool(obj.get('/Interpolate', obj.get('/I', False))),
                          masks=spatial_masks,
                          mask_resolution_known=known_masks,
                          meets_target=limiting+1e-6 >= target_ppi and known_masks,
                          required_width_px=max(1, math.ceil(width_in*target_ppi-1e-9)),
                          required_height_px=max(1, math.ceil(height_in*target_ppi-1e-9)),
                          max_scale_at_target=limiting/target_ppi if known_masks else None,
                          max_width_mm_at_target=width_in*25.4*limiting/target_ppi if known_masks else None,
                          max_height_mm_at_target=height_in*25.4*limiting/target_ppi if known_masks else None,
                          remedy='Use genuinely higher-resolution original pixels or a vector source, or reduce physical size. Interpolation does not add detail.')
            # Sheared grids may require more than the per-axis pixel minimum.
            uniform_pixel_multiplier = max(1.0, target_ppi/density['directional_min_ppi'])
            record['required_pixels_uniform_for_all_directions'] = [math.ceil(w*uniform_pixel_multiplier-1e-9), math.ceil(h*uniform_pixel_multiplier-1e-9)]
            if spatial_masks:
                record['mask_below_target'] = any(m['min_ppi']+1e-6 < target_ppi for m in spatial_masks)
            if (mask_unknown or gs_mask) and is_visible:
                issue(f'{where}/{name}: transparency group mask resolution cannot be reduced to this image sample grid', True)
            page_images.append(record)

        def walk(stream, resources, ctm, clip, clip_exact=True, alpha=1.0, gs_mask=False, active_forms=(), where='page', depth=0):
            if stream is None:
                return
            if depth > 32:
                issue(f'{where}: nested Form depth limit reached', True)
                return
            resources = _obj(resources) if resources else {}
            font_resources = _obj(resources.get('/Font', {}))
            for name, ref in font_resources.items():
                info = _font_info(ref, str(name))
                page_fonts[info['name']] = info
                if info['subtype'] == '/Type3':
                    issue(f'{where}: Type3 font glyph programs are not audited for raster images', True)
            try:
                operations = ContentStream(stream, reader).operations
            except Exception as exc:
                issue(f'{where}: cannot parse content: {type(exc).__name__}: {exc}', True)
                return
            stack, pathpoints = [], []
            pendingclip, rectangular = False, False
            xobjects = _obj(resources.get('/XObject', {}))
            states = _obj(resources.get('/ExtGState', {}))
            for op_index, (operands, op) in enumerate(operations):
                if op == b'q':
                    stack.append((ctm, clip[:], clip_exact, alpha, gs_mask))
                elif op == b'Q':
                    if stack:
                        ctm, clip, clip_exact, alpha, gs_mask = stack.pop()
                    else:
                        issue(f'{where}: unmatched graphics-state restore', True)
                elif op == b'cm':
                    ctm = _matrix(ctm, tuple(float(x) for x in operands))
                elif op == b'gs':
                    state = _obj(states.get(operands[0], {}))
                    if '/ca' in state:
                        alpha = min(1.0, max(0.0, float(state['/ca'])))
                    if '/SMask' in state:
                        gs_mask = state['/SMask'] != '/None'
                elif op in (b'm', b'l', b'c', b'v', b'y', b're'):
                    page_counts['path_ops'] += 1
                    if op == b're':
                        x, y, w, h = map(float, operands)
                        rect = [_point(ctm, a, b) for a, b in ((x,y),(x+w,y),(x+w,y+h),(x,y+h))]
                        rectangular = not pathpoints and ((abs(ctm[1]) < 1e-10 and abs(ctm[2]) < 1e-10) or (abs(ctm[0]) < 1e-10 and abs(ctm[3]) < 1e-10))
                        pathpoints.extend(rect)
                    else:
                        rectangular = False
                        pathpoints.extend(_point(ctm,float(operands[i]),float(operands[i+1])) for i in range(0,len(operands),2))
                elif op in (b'W', b'W*'):
                    pendingclip = True
                elif op in (b'n', b'S', b's', b'f', b'F', b'f*', b'B', b'B*', b'b', b'b*'):
                    if op != b'n' and pathpoints:
                        page_counts['path_paint_ops'] += 1
                    if pendingclip:
                        if pathpoints:
                            pathbox = [min(x for x,y in pathpoints), min(y for x,y in pathpoints), max(x for x,y in pathpoints), max(y for x,y in pathpoints)]
                            clip = _intersect(clip, pathbox)
                            clip_exact = clip_exact and rectangular
                        else:
                            clip = [0.0, 0.0, 0.0, 0.0]
                    pathpoints, pendingclip, rectangular = [], False, False
                elif op in (b'Tj', b'TJ', b"'", b'"'):
                    page_counts['text_draw_ops'] += 1
                elif op in (b'scn', b'SCN'):
                    patterns = _obj(resources.get('/Pattern', {}))
                    for operand in operands:
                        if isinstance(operand, str) and operand in patterns:
                            pattern = _obj(patterns[operand])
                            if pattern.get('/PatternType') == 1:
                                issue(f'{where}: tiling pattern {operand} content is not audited for raster images', True)
                elif op == b'Do':
                    name = operands[0]
                    ref = xobjects.get(name)
                    if ref is None:
                        issue(f'{where}: missing XObject {name}', True)
                        continue
                    obj = _obj(ref)
                    key = (getattr(ref, 'idnum', None), getattr(ref, 'generation', None)) if hasattr(ref, 'idnum') else ('direct', id(obj))
                    xref = getattr(ref, 'idnum', None)
                    if obj.get('/Subtype') == '/Image':
                        add_image(obj, ctm, clip, clip_exact, alpha, gs_mask, where, str(name), xref)
                    elif obj.get('/Subtype') == '/Form':
                        if key in active_forms:
                            issue(f'{where}: recursive Form cycle at {name}', True)
                            continue
                        newctm = _matrix(ctm, tuple(float(x) for x in obj.get('/Matrix', IDENTITY)))
                        formclip = clip[:]
                        form_exact = clip_exact
                        if '/BBox' in obj:
                            formclip = _intersect(clip, _bounds(newctm, [float(x) for x in obj['/BBox']]))
                            form_exact = form_exact and ((abs(newctm[1])<1e-10 and abs(newctm[2])<1e-10) or (abs(newctm[0])<1e-10 and abs(newctm[3])<1e-10))
                        else:
                            issue(f'{where}/{name}: Form BBox missing; clip visibility is uncertain')
                            form_exact = False
                        walk(obj, obj.get('/Resources', resources), newctm, formclip, form_exact, alpha, gs_mask, active_forms+(key,), where+'/'+str(name), depth+1)
                    else:
                        issue(f'{where}: unsupported XObject subtype {obj.get("/Subtype")} at {name}', True)
                elif op == b'INLINE IMAGE':
                    settings = operands.get('settings', {})
                    add_image(settings, ctm, clip, clip_exact, alpha, gs_mask, where, f'inline-{op_index}', None, inline=True)
            if stack:
                issue(f'{where}: unbalanced saved graphics states', True)

        walk(page.get('/Contents'), page.get('/Resources'), IDENTITY, initial_clip)
        images.extend(page_images)
        for k, v in page_counts.items():
            counters[k] += v
        fonts.update(page_fonts)
        issues.extend(f'Page {page_number}: {x}' for x in page_issues)
        w = (cropbox[2]-cropbox[0])*userunit
        h = (cropbox[3]-cropbox[1])*userunit
        if int(page.get('/Rotate', 0)) % 180:
            w, h = h, w
        pages.append(dict(page=page_number, width_pt=w, height_pt=h,
                          width_mm=w*25.4/72, height_mm=h*25.4/72,
                          vector=page_counts, fonts=page_fonts,
                          placement_count=len(page_images), images=page_images, issues=page_issues))

    visible = [im for im in images if im['potentially_visible']]
    main = [im for im in visible if not im['small_element']]
    below = [im for im in visible if not im['meets_target']]
    minimum = min((im['min_ppi'] for im in visible), default=None)
    primary_minimum = min((im['min_ppi'] for im in main), default=None)
    if not complete:
        status = 'needs_review'
    elif below:
        status = 'below_target'
    elif not visible:
        status = 'no_raster_images'
    else:
        status = 'meets_target'
    return dict(schema_version=1, pdf=source.name, target_ppi=target_ppi,
                status=status, min_ppi=minimum, min_large_ppi=primary_minimum,
                image_count=len(visible), placement_count=len(images),
                main_image_count=len(main), small_image_count=len(visible)-len(main),
                below_target_count=len(below), below_target_images=below,
                meets_target=not below and complete, assessment_complete=complete,
                contains_vector_paths=counters['path_paint_ops']>0,
                contains_native_text=counters['text_draw_ops']>0,
                vector=counters, fonts=fonts, pages=pages, images=images, issues=issues,
                max_figure_scale_at_target=minimum/target_ppi if minimum is not None and complete else None,
                caveats=[
                    'A vector PDF has no single DPI. Raster PPI is computed from actual pixel dimensions and physical placement, not nominal DPI metadata.',
                    'Visibility is conservative: page/CropBox, clipping-path bounds and zero opacity are considered; arbitrary path interiors, optional-content visibility and overpainting/occlusion are not fully resolved.',
                    'Attached spatial masks are included in the parent image quality limit and are not counted as independent pictures. External transparency-group masks cause needs_review.',
                    'Encountered tiling patterns, Type3 font glyph programs and annotation appearances require review because their internal raster image programs are not audited.',
                    'Pixel counts alone do not prove source detail. Consult the preservation provenance to distinguish original pixels from any previously upsampled input.',
                    'Required pixel sizes refer to genuinely better source data. Resizing a bitmap by interpolation does not satisfy the detail requirement.'
                ])


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pdf', type=Path)
    parser.add_argument('--target-ppi', type=float, default=600)
    parser.add_argument('--json-output', type=Path)
    args = parser.parse_args()
    report = assess_pdf(args.pdf, args.target_ppi)
    payload = json.dumps(report, ensure_ascii=False, indent=2)
    if args.json_output:
        # Do not overwrite a user report by accident.
        with args.json_output.open('x', encoding='utf-8') as stream:
            stream.write(payload)
    else:
        print(payload)
