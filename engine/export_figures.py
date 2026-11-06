#!/usr/bin/env python3
"""Inspect PowerPoint paper figures, render natively, and crop vector PDFs.

Python 3.10+, Windows desktop PowerPoint. Input is never saved or edited.
All geometry uses PDF points (72 pt/in); overrides are JSON, not executable code.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import html
import json
import math
import re
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

from pptx import Presentation
from pptx.shapes.shapetree import SlideShapeFactory
from PIL import Image, ImageChops, ImageDraw

EMU = 12700.0
LABEL = re.compile(r"^\s*(fig(?:ure)?\.?|table|图|表)\s*(S?\d+|[IVX]+)\b", re.I)
WIDTH_FULL = re.compile(r"full[\s_-]*width|double[\s_-]*column|two[\s_-]*column|双栏|通栏", re.I)
WIDTH_HALF = re.compile(r"half[\s_-]*width|single[\s_-]*column|one[\s_-]*column|单栏|半宽", re.I)


def sha256(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


def write_json(path, value):
    Path(path).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')


def box_union(boxes):
    return [min(b[0] for b in boxes), min(b[1] for b in boxes),
            max(b[2] for b in boxes), max(b[3] for b in boxes)]


def width_hint(text):
    a, b = bool(WIDTH_FULL.search(text)), bool(WIDTH_HALF.search(text))
    return 'full' if a and not b else 'half' if b and not a else None


def shape_info(shape):
    x, y, w, h = [float(v) / EMU for v in (shape.left, shape.top, shape.width, shape.height)]
    angle = float(shape.rotation) % 360
    if angle:
        a = math.radians(angle)
        rw, rh = abs(w * math.cos(a)) + abs(h * math.sin(a)), abs(w * math.sin(a)) + abs(h * math.cos(a))
        x, y, w, h = x + (w - rw) / 2, y + (h - rh) / 2, rw, rh
    nodes = list(shape.element.iter())
    text = '\n'.join(n.text or '' for n in nodes if n.tag.endswith('}t')).strip()
    progids = [n.get('progId', '') for n in nodes if n.tag.endswith('}oleObj')]
    has_table = any(n.tag.endswith('}tbl') for n in nodes)
    has_picture = any(n.tag.endswith('}pic') for n in nodes)
    has_chart = any(n.tag.endswith('}chart') for n in nodes)
    has_ole = bool(progids)
    # Count diagram primitives, excluding ordinary text boxes.
    primitives = sum(1 for n in nodes if n.tag.endswith('}prstGeom') and n.get('prst') not in ('rect',))
    has_drawing = primitives > 0 or any(n.tag.endswith('}cxnSp') or n.tag.endswith('}custGeom') for n in nodes)
    group = int(shape.shape_type) == 6
    if group and not has_picture and not has_ole and not has_chart:
        has_drawing = has_drawing or sum(n.tag.endswith('}sp') for n in nodes) >= 2
    if not text and not has_picture and not has_table and not has_ole:
        has_drawing = True
    hidden = any(n.tag.endswith('}cNvPr') and n.get('id') == str(shape.shape_id)
                 and n.get('hidden') in ('1', 'true') for n in nodes)
    placeholder = str(shape.placeholder_format.type) if shape.is_placeholder else None
    role = None
    m = LABEL.match(text)
    if m and not (group or has_drawing or has_table or has_picture or has_ole or has_chart):
        role = 'table_label' if m.group(1).lower() in ('table', '表') else 'figure_label'
    elif placeholder and any(k in placeholder for k in ('TITLE', 'DATE', 'FOOTER', 'SLIDE_NUMBER')):
        role = 'decoration'
    elif width_hint(text) and len(text) < 65 and not (group or has_drawing or has_picture or has_ole or has_chart):
        role = 'width_label'
    return dict(id=int(shape.shape_id), name=shape.name, type=str(shape.shape_type),
                bbox_pt=[round(v, 4) for v in (x, y, x + w, y + h)], rotation=angle,
                text=text[:2500], table=has_table, picture=has_picture, chart=has_chart,
                ole_progids=progids, drawing=has_drawing, group=group, hidden=hidden,
                placeholder=placeholder, role=role,
                mixed_table_group=bool(has_table and group and (has_chart or has_ole or any(n.tag.endswith('}sp') or n.tag.endswith('}pic') or n.tag.endswith('}cxnSp') for n in nodes))))


def find_detached_guides(shapes, margin_pt=6.0, touch_pt=1.0, overshoot_pt=24.0):
    """Return IDs of alignment/guide lines left on a slide.

    Kept, a guide would enlarge the crop and appear in the figure. A top-level
    straight line is a guide when either
      1. it touches no other object and extends more than margin_pt beyond
         the bounding box of the remaining content, or
      2. it is horizontal or vertical, overshoots the remaining content by
         more than overshoot_pt at BOTH ends (typical slide-edge guide), and
         touches no other line - only the bounding boxes of pictures, text
         boxes or shapes, whose blank margins can overlap a guide.
    Lines inside the artwork (arrows, axis ticks, leaders, separators) touch
    other objects or lie within the content and are never matched.
    """
    pool = [s for s in shapes if not s['hidden'] and not s['role'] and not s['table']]
    touch = lambda a, b: not (a[2] < b[0]-touch_pt or b[2] < a[0]-touch_pt or a[3] < b[1]-touch_pt or b[3] < a[1]-touch_pt)
    is_line = lambda s: s['type'].startswith('LINE') and not s['group']
    isolated = [s for s in pool if is_line(s)
                and not any(touch(s['bbox_pt'], o['bbox_pt']) for o in pool if o['id'] != s['id'])]
    content = [s for s in pool if s not in isolated]
    if len(content) < 3:
        return []
    guides = set()
    # Rule 2: long axis-aligned lines that overshoot everything else at both ends.
    for s in content:
        b = s['bbox_pt']
        vertical, horizontal = b[2] - b[0] < .5, b[3] - b[1] < .5
        if not is_line(s) or vertical == horizontal:
            continue
        rest = [o for o in content if o['id'] != s['id']]
        if len(rest) < 3 or any(is_line(o) and touch(b, o['bbox_pt']) for o in rest):
            continue
        u = box_union([o['bbox_pt'] for o in rest])
        low, high = (u[1] - b[1], b[3] - u[3]) if vertical else (u[0] - b[0], b[2] - u[2])
        if min(low, high) > overshoot_pt:
            guides.add(s['id'])
    # Rule 1: isolated lines outside the remaining content.
    remaining = [s for s in content if s['id'] not in guides]
    if len(remaining) >= 3:
        u = box_union([s['bbox_pt'] for s in remaining])
        for s in isolated:
            b = s['bbox_pt']
            if max(u[0]-b[0], u[1]-b[1], b[2]-u[2], b[3]-u[3]) > margin_pt:
                guides.add(s['id'])
    return [s['id'] for s in shapes if s['id'] in guides]


def graphical(s):
    return s['picture'] or s['chart'] or s['drawing'] or bool(s['ole_progids'])


def all_slide_shapes(slide):
    """Include Office math/modern objects wrapped in mc:AlternateContent.

    python-pptx's public shape iterator skips these wrappers. Select the first
    native Choice, or its Fallback if needed, for read-only metadata. The native
    renderer later audits every PowerPoint shape ID against this inventory.
    """
    supported = {'sp','grpSp','graphicFrame','cxnSp','pic'}
    result, unknown = [], []
    for element in slide.shapes._spTree:
        local = element.tag.rsplit('}',1)[-1]
        if local in ('nvGrpSpPr','grpSpPr','extLst'):
            continue
        native = [element]
        alternate = local == 'AlternateContent'
        if alternate:
            native = []
            for branch in element:
                choices = [n for n in branch if n.tag.rsplit('}',1)[-1] in supported]
                if choices:
                    native = choices
                    break
        if not native:
            unknown.append(local)
        for node in native:
            if node.tag.rsplit('}',1)[-1] not in supported:
                unknown.append(node.tag.rsplit('}',1)[-1]); continue
            try:
                info = shape_info(SlideShapeFactory(node, slide.shapes))
                info['alternate_content'] = alternate
                result.append(info)
            except Exception as exc:
                unknown.append(f'{node.tag}: {type(exc).__name__}: {exc}')
    if len({s['id'] for s in result}) != len(result):
        raise ValueError('Duplicate top-level shape IDs in source XML; inspect/repair the source copy in PowerPoint.')
    return result, unknown


def analyze(source, config, args):
    prs = Presentation(source)
    sw, sh = float(prs.slide_width) / EMU, float(prs.slide_height) / EMU
    manifest = dict(version=1, source=source.name, source_sha256=sha256(source),
                    slide_width_pt=sw, slide_height_pt=sh, slides=[], candidates=[],
                    sizing=dict(full_width_mm=args.full_width_mm, half_width_mm=args.half_width_mm,
                                provenance=config.get('width_provenance', 'User CLI/config dimensions' if args.full_width_mm else 'Preserve cropped source size; no physical paper dimensions assumed')),
                    warnings=[
                        'Width classification is inferred from explicit labels, overrides, then relative figure geometry; it is not a journal-template detector.',
                        'Unlabelled content without a native chart or Origin object is held for review. Rasterized and drawn tables cannot be classified reliably without labels or overrides.',
                        'Unlabelled panels on a slide stay together; use explicit figure labels or groups overrides to split multiple independent figures.',
                        'Only saved source contents are used. OLE uses its saved preview; fonts/effects depend on installed PowerPoint.'])
    slide_configs = config.get('slides', {})
    for si, slide in enumerate(prs.slides, 1):
        shapes, unsupported = all_slide_shapes(slide)
        hidden = slide.element.get('show') in ('0', 'false')
        scfg = slide_configs.get(str(si), {})
        record = dict(slide=si, hidden=hidden, shapes=shapes, warnings=[], unsupported_objects=unsupported)
        if unsupported:
            record['warnings'].append('Unsupported XML objects require review: ' + '; '.join(unsupported))
        guides = [] if getattr(args, 'keep_guide_lines', False) else find_detached_guides(shapes)
        for s in shapes:
            if s['id'] in guides:
                s['role'] = 'detached_guide_line'
        if guides:
            record['warnings'].append(f'Excluded detached guide line(s) {guides}: straight lines that touch no other object and lie outside the figure. '
                                      'Use --keep-guide-lines (or list them in a groups override) to keep them.')
        manifest['slides'].append(record)
        candidates = []
        active = [s for s in shapes if not s['hidden'] and not s['role']]
        labels = [s for s in shapes if not s['hidden'] and s['role'] in ('figure_label', 'table_label')]
        context = '\n'.join(s['text'] + ' ' + s['name'] for s in shapes if not s['hidden'] and s['role'])
        # Every native table receives a visible skipped entry, including tables inside a group.
        table_shapes = [s for s in active if s['table']]
        for s in table_shapes:
            mixed = s['mixed_table_group']
            candidates.append(dict(kind='review' if mixed else 'table', shape_ids=[s['id']],
                                   label='Mixed group with table' if mixed else 'Native table',
                                   reason='native_table_in_mixed_group' if mixed else 'native_table', width=None))
        active = [s for s in active if not s['table']]
        if labels:
            buckets = [[] for _ in labels]
            centers = [((s['bbox_pt'][0] + s['bbox_pt'][2]) / 2,
                        (s['bbox_pt'][1] + s['bbox_pt'][3]) / 2) for s in labels]
            horizontal = len(labels) > 1 and max(v[0] for v in centers) - min(v[0] for v in centers) > sw * .25 and max(v[1] for v in centers) - min(v[1] for v in centers) < sh * .20
            ambiguous_labels = len(labels) > 1 and not horizontal
            boundaries = [(a[0]+b[0])/2 for a,b in zip(sorted(centers),sorted(centers)[1:])] if horizontal else []
            for s in active:
                b = s['bbox_pt']; cx, cy = (b[0] + b[2]) / 2, (b[1] + b[3]) / 2
                if any(b[0] < edge-2 and b[2] > edge+2 for edge in boundaries):
                    ambiguous_labels = True
                distances = [abs(cx - v[0]) / sw if horizontal else abs(cy - v[1]) / sh + .2 * abs(cx - v[0]) / sw for v in centers]
                buckets[min(range(len(labels)), key=lambda k: distances[k])].append(s)
            for label, members in zip(labels, buckets):
                if not members:
                    record['warnings'].append('Label has no separate content: ' + label['text'][:100])
                    ambiguous_labels = True
                    continue
                kind = 'table' if label['role'] == 'table_label' else 'figure'
                if kind == 'figure' and not any(graphical(s) for s in members):
                    kind = 'review'
                candidates.append(dict(kind=kind, shape_ids=[s['id'] for s in members],
                                       label=label['text'], reason='explicit_' + label['role'],
                                       width=width_hint(label['text']) or width_hint(context)))
            if ambiguous_labels:
                record['warnings'].append('Multiple captions have ambiguous vertical/crossing/empty assignments. Review grouping before export.')
                for c in candidates:
                    if c['reason'].startswith('explicit_'):
                        c['kind'] = 'review'; c['reason'] = 'ambiguous_multiple_labels'
        elif active:
            graphics = [s for s in active if graphical(s)]
            if graphics:
                has_reliable = all(s['chart'] or any('origin' in p.lower() for p in s['ole_progids']) for s in graphics)
                spreadsheet = any(any('excel' in p.lower() or 'word' in p.lower() for p in s['ole_progids']) for s in graphics)
                kind = 'figure' if has_reliable and not spreadsheet else 'review'
                candidates.append(dict(kind=kind, shape_ids=[s['id'] for s in active],
                                       label='Unlabelled slide content', width=width_hint(context),
                                       reason='native_chart_or_Origin' if kind == 'figure' else 'unlabelled_content_requires_review'))
            else:
                candidates.append(dict(kind='skip', shape_ids=[s['id'] for s in active], label='Text-only slide', width=None, reason='no_graphical_content'))
        # An explicit groups list replaces automatic grouping for this slide.
        if 'groups' in scfg:
            candidates = []
            used = set()
            for j, group in enumerate(scfg['groups'], 1):
                ids = group.get('shape_ids')
                if ids == 'all_content':
                    ids = [s['id'] for s in shapes if not s['hidden'] and not s['role']]
                if not isinstance(ids, list) or not ids or not all(type(v) is int for v in ids):
                    raise ValueError(f'Slide {si} group {j}: shape_ids must be a nonempty integer list or all_content')
                if len(set(ids)) != len(ids) or used.intersection(ids):
                    raise ValueError(f'Slide {si}: duplicate shape assignment in groups')
                used.update(ids)
                candidates.append(dict(kind=group.get('kind', 'figure'), shape_ids=ids,
                                       label=group.get('label', f'Manual group {j}'), width=group.get('width'),
                                       reason='manual_override', crop_pt=group.get('crop_pt'),
                                       allow_offslide=group.get('allow_offslide', scfg.get('allow_offslide', False))))
        ids_map = {s['id']: s for s in shapes}
        assigned = set()
        for c in candidates:
            if c['kind'] not in ('figure', 'table', 'review', 'skip'):
                raise ValueError(f'Slide {si}: invalid kind {c["kind"]!r}')
            if not set(c['shape_ids']) <= ids_map.keys():
                raise ValueError(f'Slide {si}: unknown shape_ids {set(c["shape_ids"]) - ids_map.keys()}')
            c.update(slide=si, bbox_pt=box_union([ids_map[i]['bbox_pt'] for i in c['shape_ids']]), warnings=[])
            if unsupported and c['kind']=='figure':
                c['kind']='review'; c['reason']='unsupported_XML_objects'
                c['warnings'].append('Some slide objects could not be inventoried; export is held to prevent silent omissions.')
            assigned.update(c['shape_ids'])
            if scfg.get('kind') == 'skip':
                c['kind'] = 'skip'; c['reason'] = 'manual_slide_skip'
            if scfg.get('width') and c['kind'] == 'figure':
                c['width'] = scfg['width']; c['width_reason'] = 'manual_slide_override'
            if hidden and not args.include_hidden:
                c['original_kind'] = c['kind']; c['kind'] = 'skip'; c['reason'] = 'hidden_slide'
            b = c['bbox_pt']
            canvas = [min(0, b[0]-4), min(0, b[1]-4), max(sw, b[2]+4), max(sh, b[3]+4)]
            if b[0] < -.2 or b[1] < -.2 or b[2] > sw + .2 or b[3] > sh + .2:
                # An explicit per-figure/per-slide allow_offslide keeps the old "expand" behaviour.
                mode = 'expand' if c.get('allow_offslide', scfg.get('allow_offslide')) else getattr(args, 'offslide', 'review')
                c['offslide_mode'] = mode
                if mode == 'clip':
                    clipped = [max(0, b[0]), max(0, b[1]), min(sw, b[2]), min(sh, b[3])]
                    canvas = [0, 0, sw, sh]
                    c['source_bbox_pt'] = b
                    if clipped[0] >= clipped[2] or clipped[1] >= clipped[3]:
                        c['warnings'].append('All selected content lies outside the slide; nothing would remain after clipping at the slide edge.')
                        if c['kind'] == 'figure':
                            c['kind'] = 'review'; c['reason'] = 'content_entirely_off_slide'
                    else:
                        c['bbox_pt'] = clipped
                        c['warnings'].append('Parts of some objects lie outside the slide (often transparent text-box or picture margins). '
                                             'They are clipped at the slide edge, exactly as PowerPoint shows the slide.')
                elif mode == 'expand':
                    c['warnings'].append('Selected content extends outside the original slide; it is rendered on an enlarged temporary canvas to retain it.')
                else:
                    c['warnings'].append('Selected content extends outside the original slide. Use --offslide clip (cut at the slide edge, '
                                         'as PowerPoint shows it) or --offslide expand (keep it on an enlarged canvas).')
                    if c['kind'] == 'figure':
                        c['kind'] = 'review'; c['reason'] = 'off_slide_content'
            c['canvas_pt'] = canvas
            if any(ids_map[i]['ole_progids'] for i in c['shape_ids']):
                c['warnings'].append('OLE cached appearance only; inspect exported preview for missing content.')
            if any(ids_map[i]['table'] for i in c['shape_ids']) and c['kind'] == 'figure':
                c['warnings'].append('Manual figure override includes a native table.')
            if c.get('width') not in (None, 'full', 'half'):
                raise ValueError(f'Slide {si}: width must be full or half')
            if c.get('crop_pt') is not None:
                rect = c['crop_pt']
                if not isinstance(rect,list) or len(rect)!=4 or not all(type(v) in (int,float) and math.isfinite(v) for v in rect) or rect[0]>=rect[2] or rect[1]>=rect[3]:
                    raise ValueError(f'Slide {si}: crop_pt must be [left, top, right, bottom] in original slide points')
        candidates.sort(key=lambda c: (round(c['bbox_pt'][1] / 12), c['bbox_pt'][0]))
        for j, c in enumerate(candidates, 1):
            c['id'] = f's{si}f{j}'
        record['unassigned_shapes'] = [dict(id=s['id'], reason=s['role'] or ('hidden_shape' if s['hidden'] else 'not_selected_by_override')) for s in shapes if s['id'] not in assigned]
        manifest['candidates'].extend(candidates)
    unknown = set(slide_configs) - {str(i) for i in range(1, len(prs.slides) + 1)}
    if unknown:
        raise ValueError(f'Config references nonexistent slides: {sorted(unknown)}')
    candidates = manifest['candidates']
    reference = max([c['bbox_pt'][2] - c['bbox_pt'][0] for c in candidates if c['kind'] in ('figure', 'review')] or [sw])
    manifest['width_reference_pt'] = reference
    for c in candidates:
        if c['kind'] in ('figure', 'review'):
            if not c.get('width'):
                ratio = (c['bbox_pt'][2] - c['bbox_pt'][0]) / reference
                c['width'] = 'half' if ratio < args.half_threshold else 'full'
                c['width_reason'] = f'geometry ratio {ratio:.3f} of widest candidate, threshold {args.half_threshold:.3f}'
                c['warnings'].append('Width is inferred from geometry; verify against intended manuscript layout.')
            else:
                c.setdefault('width_reason', 'explicit label or config override')
    return manifest


def measure_crop(source, candidate, args, bounds_pdf=None):
    """Find the crop rectangle (points, top-left origin, rendered page) of visible content."""
    import pypdfium2 as pdfium
    from pypdf import PdfReader
    reader = PdfReader(str(source))
    if len(reader.pages) != 1:
        raise ValueError(f'Expected exactly one rendered slide: {source}')
    pw, ph = float(reader.pages[0].mediabox.width), float(reader.pages[0].mediabox.height)
    # Use the same native-rendered bounds after pixel restoration. Sharper
    # alpha edges must not silently change the published physical layout.
    doc = pdfium.PdfDocument(str(bounds_pdf or source))
    bitmap = doc[0].render(scale=args.trim_dpi / 72)
    im = bitmap.to_pil().convert('RGB')
    bitmap.close(); doc.close()
    scale_x, scale_y = im.width / pw, im.height / ph
    origin = candidate.get('canvas_pt', [0,0,pw,ph])[:2]
    manual = candidate.get('crop_pt')
    if manual is not None:
        if len(manual) != 4 or not all(type(v) in (int, float) and math.isfinite(v) for v in manual):
            raise ValueError('crop_pt must be four finite numbers [left, top, right, bottom]')
        box = [manual[0]-origin[0], manual[1]-origin[1], manual[2]-origin[0], manual[3]-origin[1]]
        candidate['crop_method'] = 'manual_rectangle'
    else:
        # Detect visible ink on the isolated white slide; final PDF stays vector.
        corners = [im.getpixel(p) for p in ((0, 0), (im.width-1, 0), (0, im.height-1), (im.width-1, im.height-1))]
        white = all(min(rgb) > 246 for rgb in corners)
        if white:
            diff = ImageChops.difference(im, Image.new('RGB', im.size, 'white')).convert('L')
            ink = diff.point(lambda v: 255 if v > 6 else 0).getbbox()
            if ink is None:
                raise ValueError(f'Blank rendering for {candidate["id"]}; no figure PDF produced')
            box = [ink[0] / scale_x, ink[1] / scale_y, ink[2] / scale_x, ink[3] / scale_y]
            candidate['crop_method'] = f'visible_ink_at_{args.trim_dpi}_dpi'
            if ink[0] == 0 or ink[1] == 0 or ink[2] == im.width or ink[3] == im.height:
                candidate['warnings'].append('Visible ink touches slide edge; check original for off-slide clipping.')
        else:
            b = candidate['bbox_pt']
            box = [b[0]-origin[0], b[1]-origin[1], b[2]-origin[0], b[3]-origin[1]]
            candidate['crop_method'] = 'geometry_nonwhite_background'
            candidate['warnings'].append('Nonwhite background: using shape geometry instead of white-space trimming.')
        pad = args.padding_pt
        box = [max(0, box[0]-pad), max(0, box[1]-pad), min(pw, box[2]+pad), min(ph, box[3]+pad)]
    x0, top, x1, bottom = box
    if not 0 <= x0 < x1 <= pw or not 0 <= top < bottom <= ph:
        raise ValueError(f'Invalid crop rectangle for {candidate["id"]}: {box}')
    candidate['_crop'] = dict(box=list(box), page_pt=[pw, ph], origin=list(origin))
    return candidate['_crop']


def layout_figure(candidate, args, reference_scale=None):
    """Choose the content scale and final page width for a measured figure.

    fill (default): the content fills the target width exactly.
    match-full: a half-width figure is drawn at the same scale as the full-width
    figures (reference_scale), so text has the same printed size, and centred on
    a page exactly half_width_mm wide. If it would not fit at that scale, it is
    reduced to fit (never enlarged beyond filling the half width).
    """
    box = candidate['_crop']['box']
    w = box[2] - box[0]
    target_mm = args.full_width_mm if candidate['width'] == 'full' else args.half_width_mm
    if not target_mm:
        candidate['content_scale'] = 1.0
        candidate['scale_reason'] = 'no target width: source size kept'
        return 1.0, w
    target_pt = target_mm / 25.4 * 72
    fill = target_pt / w
    mode = getattr(args, 'half_width_scale', 'fill')
    if candidate['width'] == 'half' and mode == 'match-full' and reference_scale:
        if reference_scale <= fill:
            candidate['content_scale'] = reference_scale
            candidate['scale_reason'] = (f'same scale as the full-width figures ({reference_scale:.3f}); '
                                         f'centred with {(target_pt - w*reference_scale)/2*25.4/72:.1f} mm side margins')
            return reference_scale, target_pt
        candidate['warnings'].append('At the full-width figures\' scale this half-width figure would be wider than '
                                     f'{target_mm:g} mm, so it is reduced to fit (text slightly smaller than in full-width figures).')
    candidate['content_scale'] = fill
    candidate['scale_reason'] = f'scaled to fill {target_mm:g} mm'
    return fill, target_pt


def full_width_reference_scale(selected, args):
    """Median scale of the full-width figures (measured crops), or None."""
    if not args.full_width_mm:
        return None
    target = args.full_width_mm / 25.4 * 72
    scales = sorted(target / (c['_crop']['box'][2] - c['_crop']['box'][0])
                    for c in selected if c['width'] == 'full' and '_crop' in c)
    if not scales:
        return None
    mid = len(scales) // 2
    return scales[mid] if len(scales) % 2 else (scales[mid-1] + scales[mid]) / 2


def write_preview(output, candidate, args):
    import pypdfium2 as pdfium
    doc = pdfium.PdfDocument(str(output))
    preview_dpi = getattr(args, 'preview_dpi', 144)
    bmp = doc[0].render(scale=preview_dpi / 72)
    bmp.to_pil().convert('RGB').save(output.with_suffix('.png'))
    bmp.close(); doc.close()
    candidate['preview_dpi'] = preview_dpi


def crop_pdf(source, output, candidate, args, bounds_pdf=None, reference_scale=None, preview=True):
    """Crop the rendered page to the figure, scale/centre it, write a vector PDF."""
    from pypdf import PdfReader, PdfWriter, Transformation
    from pypdf.generic import RectangleObject
    if '_crop' not in candidate:
        measure_crop(source, candidate, args, bounds_pdf)
    crop = candidate['_crop']
    box, (pw, ph), origin = crop['box'], crop['page_pt'], crop['origin']
    reader = PdfReader(str(source))
    writer = PdfWriter()
    page = writer.add_page(reader.pages[0])
    x0, top, x1, bottom = box
    # PDF uses bottom-left origin; preserve all original vector/font resources.
    y0, y1 = ph - bottom, ph - top
    w, h = x1 - x0, y1 - y0
    scale, page_w = layout_figure(candidate, args, reference_scale)
    content_w, content_h = w * scale, h * scale
    dx = (page_w - content_w) / 2 if page_w > content_w + 1e-6 else 0.0
    page.add_transformation(Transformation().translate(-x0, -y0).scale(scale, scale).translate(dx, 0))
    for attr in ('mediabox', 'cropbox', 'trimbox', 'bleedbox', 'artbox'):
        setattr(page, attr, RectangleObject([0, 0, max(page_w, content_w), content_h]))
    page.compress_content_streams(level=9)
    from privacy import clean_pdf
    candidate['privacy_cleanup'] = clean_pdf(writer)
    writer.add_metadata({'/Title': output.stem, '/Creator': 'Paper Figure Exporter (native PowerPoint + pypdf)'})
    with open(output, 'xb') as f:
        writer.write(f)
    final = PdfReader(str(output)).pages[0]
    candidate['export_crop_pt'] = [round(box[0]+origin[0],4),round(box[1]+origin[1],4),round(box[2]+origin[0],4),round(box[3]+origin[1],4)]
    candidate['output_width_mm'] = round(float(final.mediabox.width) * 25.4 / 72, 4)
    candidate['output_height_mm'] = round(float(final.mediabox.height) * 25.4 / 72, 4)
    candidate['content_width_mm'] = round(content_w * 25.4 / 72, 4)
    candidate['side_margin_mm'] = round(dx * 25.4 / 72, 4)
    candidate['aspect_ratio_preserved'] = abs(content_w / content_h - w / h) < 1e-6
    candidate['pdf_bytes'] = output.stat().st_size
    candidate.pop('_crop', None)
    if preview:
        write_preview(output, candidate, args)


REDACT = None  # set per run: replaces absolute folder paths in everything written


def write_report(out, manifest):
    if REDACT is not None:
        manifest = REDACT(manifest)
    write_json(out / 'manifest.json', manifest)
    keys = ['id', 'slide', 'kind', 'label', 'width', 'width_reason', 'reason', 'output', 'output_width_mm', 'output_height_mm', 'warnings']
    with open(out / 'manifest.csv', 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=keys, extrasaction='ignore'); writer.writeheader()
        for c in manifest['candidates']:
            writer.writerow({**c, 'warnings': ' | '.join(c['warnings'])})
    esc = lambda v: html.escape(str(v))
    body = ['<!doctype html><meta charset="utf-8"><title>Paper figure review</title><style>body{font:16px system-ui;max-width:1180px;margin:36px auto;padding:0 20px;color:#223}img{max-width:100%;border:1px solid #ddd}section{border-top:1px solid #bbb;padding:20px 0}.review{color:#aa5300}.table,.skip{color:#667}code{white-space:pre-wrap}td,th{padding:8px;text-align:left;border-bottom:1px solid #ddd}table{border-collapse:collapse;width:100%}.thumb{max-height:460px}</style>',
            '<h1>Paper figure review</h1>', f'<p>Source: {esc(manifest["source"])}<br>SHA-256: <code>{manifest["source_sha256"]}</code></p>',
            f'<p><strong>Status: {esc(manifest.get("status", "analyzed"))}</strong> | Figures: {sum(c["kind"]=="figure" for c in manifest["candidates"])} | Tables skipped: {sum(c["kind"]=="table" for c in manifest["candidates"])} | Held for review: {sum(c["kind"]=="review" for c in manifest["candidates"])}</p>',
            f'<p>Sizing: {esc(manifest["sizing"])}</p>', '<p>Review items are held, tables are skipped. Captions, title placeholders and width labels are excluded from exported figure selections. Source previews show candidate boxes and shape IDs.</p>',
            '<ul>' + ''.join('<li>' + esc(w) + '</li>' for w in manifest['warnings']) + '</ul>']
    if any(c.get('image_quality') for c in manifest['candidates']):
        from quality_report import write_quality_report
        write_quality_report(out, manifest)
        body.append('<p><a href="quality_report.html">Image resolution and original-pixel preservation report</a> (actual PPI at final physical size; vectors have no fixed DPI).</p>')
    for slide in manifest['slides']:
        si = slide['slide']; rows = [c for c in manifest['candidates'] if c['slide'] == si]
        body.append(f'<section><h2>Slide {si}' + (' (hidden)' if slide['hidden'] else '') + '</h2>')
        preview = out / 'rendered' / f'slide-{si:03}.png'
        if preview.exists():
            im = Image.open(preview).convert('RGB'); draw = ImageDraw.Draw(im)
            sx, sy = im.width / manifest['slide_width_pt'], im.height / manifest['slide_height_pt']
            for s in slide['shapes']:
                b = s['bbox_pt']; xy = [b[0]*sx, b[1]*sy, b[2]*sx, b[3]*sy]
                draw.rectangle(xy, outline='#888888', width=1); draw.text((max(0, xy[0]), max(0, xy[1])), str(s['id']), fill='#555555', stroke_width=1, stroke_fill='white')
            for c in rows:
                b = c['bbox_pt']; color = {'figure':'#00825a','table':'#a03535','review':'#c77800','skip':'#555555'}[c['kind']]
                draw.rectangle([b[0]*sx,b[1]*sy,b[2]*sx,b[3]*sy], outline=color,width=3)
            name = f'slide-{si:03}-marked.png'; im.save(out / name)
            body.append(f'<img src="{name}" alt="Slide {si} candidate boxes and shape IDs">')
        for c in rows:
            body.append(f'<h3 class="{c["kind"]}">{esc(c["id"])}: {esc(c["kind"])} / {esc(c.get("width") or "-")}</h3><p>{esc(c["label"])}<br>Shapes: {esc(c["shape_ids"])}<br>{esc(c["reason"])}; {esc(c.get("width_reason", ""))}</p>')
            if c.get('output'):
                facts = [f'{c["output_width_mm"]} × {c["output_height_mm"]} mm']
                if c.get('content_scale') is not None:
                    facts.append(f'text scale {c["content_scale"]:.3f} (10 pt on the slide prints at {10*c["content_scale"]:.1f} pt)')
                if c.get('side_margin_mm'):
                    facts.append(f'centred, {c["side_margin_mm"]:.1f} mm side margins')
                if c.get('pdf_bytes'):
                    facts.append(f'{c["pdf_bytes"]/1048576:.2f} MB')
                body.append(f'<a href="{c["output"]}">{c["output"]}</a> — ' + esc(' · '.join(facts)) + f'<br><img class="thumb" src="{Path(c["output"]).stem}.png">')
            if c['warnings']:
                body.append('<ul>' + ''.join('<li>' + esc(w) + '</li>' for w in c['warnings']) + '</ul>')
        body.append('<details><summary>All shapes, exclusions and slide warnings</summary><pre>' + esc(json.dumps(slide, ensure_ascii=False, indent=2)) + '</pre></details></section>')
    (out / 'review.html').write_text('\n'.join(body), encoding='utf-8')
    # A separate, directly editable configuration can be reused with --config.
    overrides = dict(source_sha256=manifest['source_sha256'], width_provenance=manifest['sizing']['provenance'], full_width_mm=manifest['sizing']['full_width_mm'],
                     half_width_mm=manifest['sizing']['half_width_mm'], slides={})
    for si in {c['slide'] for c in manifest['candidates']}:
        overrides['slides'][str(si)] = {'groups': [{k: c[k] for k in ('kind','shape_ids','label','width','allow_offslide','crop_pt') if c.get(k) is not None} for c in manifest['candidates'] if c['slide'] == si]}
    write_json(out / 'overrides.generated.json', overrides)


def parse_args(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('source', type=Path, help='Saved .pptx/.pptm file (never modified)')
    p.add_argument('--output', type=Path, help='New or empty output folder; existing files are never overwritten')
    p.add_argument('--config', type=Path, help='JSON grouping/kind/width overrides')
    p.add_argument('--analyze-only', action='store_true', help='Write JSON/CSV/HTML inventory without calling PowerPoint')
    p.add_argument('--full-width-mm', type=float, help='Target full width including padding; default preserves source size')
    p.add_argument('--half-width-mm', type=float, help='Target half width including padding; if omitted use full/2')
    p.add_argument('--half-threshold', type=float, default=.67, help='Half-width geometric ratio threshold (default .67)')
    p.add_argument('--padding-pt', type=float, default=2, help='Whitespace around visible ink (default 2 pt before scaling)')
    p.add_argument('--trim-dpi', type=int, default=216, help='Resolution for bounding-box detection only; PDF is not rasterized')
    p.add_argument('--min-ppi', type=float, default=600, help='Report raster elements below this actual final-size PPI (default 600); never upsamples')
    p.add_argument('--preview-dpi', type=int, default=600, help='PNG preview rendering resolution (default 600); does not change the PDF or add source detail')
    p.add_argument('--strict-ppi', action='store_true', help='Return exit code 3 if the quality target is unmet or incomplete; retain outputs and report')
    p.add_argument('--native-images', action='store_true', help='Diagnostic fallback: retain native PowerPoint PDF images without original-pixel restoration')
    p.add_argument('--include-hidden', action='store_true', help='Include hidden slides (otherwise listed as skipped)')
    p.add_argument('--offslide', choices=('review', 'clip', 'expand'), default='review',
                   help='Figures whose objects cross the slide edge: review = hold for review (default); '
                        'clip = cut at the slide edge exactly as PowerPoint shows the slide; expand = keep on an enlarged canvas')
    p.add_argument('--half-width-scale', choices=('fill', 'match-full'), default='fill',
                   help='fill: half-width figures fill the half width (default); match-full: draw them at the same scale as '
                        'the full-width figures (same printed text size) and centre them on the half-width page')
    p.add_argument('--max-file-mb', type=float, help='Keep each figure PDF at or below this size: lossless re-encoding first, then '
                   'reduce only pictures above --min-ppi, never below it. Never JPEG. The untouched file is kept in lossless/')
    p.add_argument('--keep-intermediate', action='store_true', help="Keep PowerPoint's raw per-figure PDFs in rendered/ "
                   '(they contain the deck metadata, e.g. author); removed after a successful run by default')
    p.add_argument('--keep-guide-lines', action='store_true', help='Do not exclude detached guide lines (lines touching nothing, outside the figure)')
    p.add_argument('--timeout', type=int, default=600, help='PowerPoint renderer timeout in seconds')
    p.add_argument('--powershell-exe', type=Path, help='Renderer shell executable; the Windows launcher passes its own PowerShell edition')
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    source = args.source.resolve(strict=True)
    if source.suffix.lower() not in ('.pptx', '.pptm', '.ppsx', '.ppsm'):
        raise ValueError('Save legacy .ppt as .pptx in PowerPoint first; no legacy conversion is done on the source.')
    config = json.loads(args.config.read_text(encoding='utf-8-sig')) if args.config else {}
    if config.get('source_sha256') and config['source_sha256'].lower() != sha256(source):
        raise ValueError('Override config belongs to a different saved PPT version. Run analyze-only again and review the new shape IDs before reusing selections.')
    for field in ('full_width_mm','half_width_mm'):
        if getattr(args, field) is None:
            setattr(args, field, config.get(field))
    if args.half_width_mm is None and args.full_width_mm is not None:
        args.half_width_mm = args.full_width_mm / 2
    if args.half_width_mm and not args.full_width_mm:
        raise ValueError('Specify full width as well as half width')
    for value in (args.full_width_mm, args.half_width_mm):
        if value is not None and (not math.isfinite(value) or value <= 0):
            raise ValueError('Widths must be finite positive millimeters')
    if not 0 < args.half_threshold < 1 or not 72 <= args.trim_dpi <= 600 or not math.isfinite(args.padding_pt) or not 0 <= args.padding_pt <= 72:
        raise ValueError('Invalid threshold, trim DPI or padding')
    if args.max_file_mb is not None and (not math.isfinite(args.max_file_mb) or args.max_file_mb <= 0):
        raise ValueError('--max-file-mb must be a positive number')
    if not math.isfinite(args.min_ppi) or args.min_ppi <= 0 or not 72 <= args.preview_dpi <= 1200:
        raise ValueError('Minimum PPI must be positive; preview DPI must be 72 through 1200')
    out = (args.output or source.parent / (source.stem + '_figures_' + datetime.now().strftime('%Y%m%d_%H%M%S_%f'))).resolve()
    if out.exists() and (not out.is_dir() or any(out.iterdir())):
        raise FileExistsError(f'Refusing to overwrite nonempty output directory: {out}')
    out.mkdir(parents=True, exist_ok=True)
    global REDACT
    from privacy import Redactor
    REDACT = Redactor.for_run(source, out)
    manifest = analyze(source, config, args)
    manifest['quality_settings'] = {'target_ppi':args.min_ppi, 'preview_dpi':args.preview_dpi,
                                    'max_file_mb':args.max_file_mb, 'half_width_scale':args.half_width_scale,
                                    'original_pixels':not args.native_images, 'strict_ppi':args.strict_ppi}
    manifest['status'] = 'analyzed'
    selected = [c for c in manifest['candidates'] if c['kind'] == 'figure']
    write_json(out / 'render_plan.json', {'figures': [{k:c[k] for k in ('id','slide','shape_ids','canvas_pt')} for c in selected],
               'known_shapes': {str(s['slide']):[v['id'] for v in s['shapes']] for s in manifest['slides']}, 'suppress_master_shapes':True})
    write_report(out, manifest)
    if not args.analyze_only:
        print(f'Rendering {len(selected)} figure(s) with PowerPoint; previewing all {len(manifest["slides"])} slide(s). Source stays unchanged.', flush=True)
        shell = str(args.powershell_exe.resolve(strict=True)) if args.powershell_exe else (shutil.which('powershell.exe') or shutil.which('pwsh'))
        if not shell:
            raise RuntimeError('Native export needs Windows PowerShell and installed Microsoft PowerPoint. Analyze-only works separately.')
        try:
            subprocess.run([shell, '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File',
                            str(Path(__file__).with_name('ppt_render.ps1')), '-Source', str(source),
                            '-Plan', str(out / 'render_plan.json'), '-OutputDir', str(out / 'rendered')],
                           check=True, timeout=args.timeout)
            # Pass 1: restore original pixels and measure every crop, so the
            # half-width layout can use the full-width figures' actual scale.
            sources = {}
            for c in selected:
                native_pdf = out / 'rendered' / (c['id'] + '.pdf')
                image_pdf = native_pdf
                if not args.native_images:
                    from preserve_images import restore_images
                    image_pdf = out / 'rendered' / (c['id'] + '-source-pixels.pdf')
                    c['image_preservation'] = restore_images(source, native_pdf, image_pdf,
                        c['slide'], c['shape_ids'], c['canvas_pt'],
                        slide_preview=out / 'rendered' / f'slide-{c["slide"]:03}.png',
                        slide_size_pt=(manifest['slide_width_pt'], manifest['slide_height_pt']))
                    fixed = sum(1 for im in c['image_preservation']['images'] if im.get('notes'))
                    if fixed:
                        c['warnings'].append(f'{fixed} picture(s): PowerPoint\'s PDF export had filled their transparent areas; '
                                             'original pixels and transparency restored to match the slide.')
                else:
                    c['image_preservation'] = {'mode':'native_images', 'issues':['Original-pixel restoration was disabled explicitly.']}
                measure_crop(image_pdf, c, args, bounds_pdf=native_pdf)
                sources[c['id']] = (image_pdf, native_pdf)
            reference_scale = full_width_reference_scale(selected, args)
            manifest['full_width_reference_scale'] = reference_scale
            # Pass 2: write each figure, apply the file-size limit, preview, assess.
            from image_quality import assess_pdf
            for i, c in enumerate(selected, 1):
                name = f'Fig{i}.pdf'
                image_pdf, native_pdf = sources[c['id']]
                crop_pdf(image_pdf, out / name, c, args, bounds_pdf=native_pdf,
                         reference_scale=reference_scale, preview=False)
                c['output'] = name
                if args.max_file_mb:
                    from size_control import fit_file_size, MB
                    try:
                        c['file_size'] = fit_file_size(out / name, int(args.max_file_mb * MB), args.min_ppi,
                                                       keep_copy=out / 'lossless' / name)
                    except Exception as exc:  # never lose a finished figure over size control
                        size = (out / name).stat().st_size
                        c['file_size'] = dict(changed=False, meets_limit=size <= args.max_file_mb * MB, original_bytes=size,
                                              final_bytes=size, notes=[f'Size control could not run: {type(exc).__name__}: {exc}'])
                    c['pdf_bytes'] = (out / name).stat().st_size
                    fs = c['file_size']
                    if fs['changed']:
                        how = ('lossless re-encoding' if fs['method'] == 'lossless' else
                               f'lossless re-encoding + pictures above {fs["ppi_cap"]:g} PPI reduced to it')
                        c['warnings'].append(f'File size {fs["original_bytes"]/MB:.1f} MB -> {fs["final_bytes"]/MB:.1f} MB ({how}). '
                                             f'Untouched version: lossless/{name}.')
                    if not fs['meets_limit']:
                        c['warnings'].append(f'Still above the {args.max_file_mb:g} MB limit. ' + ' '.join(fs['notes']))
                write_preview(out / name, c, args)
                c['image_quality'] = assess_pdf(out / name, target_ppi=args.min_ppi)
                if not c['image_quality']['meets_target'] or not c['image_quality']['assessment_complete']:
                    c['warnings'].append(f'Raster quality target {args.min_ppi:g} PPI is unmet or incomplete. See quality_report.html for actual pixels, masks and remedies.')
                size_note = f', {c["pdf_bytes"]/1048576:.1f} MB'
                if c.get('file_size', {}).get('changed'):
                    size_note += f' (was {c["file_size"]["original_bytes"]/1048576:.1f} MB)'
                if c.get('file_size') and not c['file_size']['meets_limit']:
                    size_note += f' STILL OVER {args.max_file_mb:g} MB'
                print(f'{name}: slide {c["slide"]}, {c["width"]}, {c["output_width_mm"]} x {c["output_height_mm"]} mm, '
                      f'text scale {c["content_scale"]:.2f}{size_note}', flush=True)
            manifest['status'] = 'completed_with_review' if any(c['kind']=='review' for c in manifest['candidates']) else 'completed'
        except BaseException as exc:
            manifest['status'] = 'failed'
            manifest['error'] = str(exc)
            manifest['warnings'].append('Export stopped; inspect completed entries and rendered/render_status.json. No later figures were silently skipped.')
            write_report(out, manifest)
            raise
        finally:
            for name in ('render_status.json', 'shape_inventory.json'):
                log = out / 'rendered' / name
                try:
                    if log.exists():
                        write_json(log, REDACT(json.loads(log.read_text(encoding='utf-8-sig'))))
                except Exception:
                    pass
            manifest['source_sha256_after'] = sha256(source)
            manifest['source_unchanged'] = manifest['source_sha256_after'] == manifest['source_sha256']
            write_report(out, manifest)
        if not args.keep_intermediate and manifest['status'] != 'failed':
            # PowerPoint's raw PDFs carry the deck's author/title metadata; the
            # figures are final, so remove them (slide images and logs are kept).
            for raw in (out / 'rendered').glob('*.pdf'):
                raw.unlink()
        if not manifest['source_unchanged']:
            manifest['status'] = 'failed'
            write_report(out, manifest)
            raise RuntimeError('Source changed during export. Rerun from a stable saved copy.')
        write_report(out, manifest)
    counts = {k:sum(c['kind'] == k for c in manifest['candidates']) for k in ('figure','table','review','skip')}
    print(json.dumps({'output':str(out), 'status':manifest['status'], 'counts':counts}, ensure_ascii=False))
    if counts['review']:
        print('REVIEW REQUIRED: inspect review.html, edit overrides.generated.json kind/shape_ids/width, then rerun into a NEW output folder.', file=sys.stderr)
        return 2
    quality_warning = any(c.get('image_quality') and (not c['image_quality']['meets_target'] or not c['image_quality']['assessment_complete']) for c in selected)
    if quality_warning:
        print(f'QUALITY NOTICE: some raster elements are below {args.min_ppi:g} PPI or could not be fully assessed. Open quality_report.html. No interpolation or automatic size reduction was used.', file=sys.stderr)
        if args.strict_ppi:
            return 3
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception as exc:
        print(f'ERROR: {exc}', file=sys.stderr)
        sys.exit(1)
