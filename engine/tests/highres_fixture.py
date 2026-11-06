"""Create a synthetic source-pixel preservation fixture (no Office automation).

All test rasters are generated locally. The PPTX is intentionally rich in
cropping, alpha, rotation, flips, nested groups, z-order and image effects.
"""
from __future__ import annotations

import hashlib
import json
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / 'highres_fixture_assets'
NS = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}


def label(shapes, text, x, y, w=3.4, h=.3, size=15, color=(0, 0, 0)):
    shape = shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    shape.name = text
    shape.text_frame.margin_left = 0
    shape.text_frame.margin_top = 0
    shape.text_frame.margin_right = 0
    shape.text_frame.margin_bottom = 0
    p = shape.text_frame.paragraphs[0]
    p.text = text
    p.font.name = 'Arial'
    p.font.size = Pt(size)
    p.font.color.rgb = RGBColor(*color)
    return shape


def background(shapes, x, y, w, h, rgb):
    s = shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = RGBColor(*rgb)
    s.line.fill.background()
    return s


def add_picture(shapes, asset, name, x, y, w, h, rotation=0, flip_h=False, flip_v=False, crop=None):
    s = shapes.add_picture(str(asset), Inches(x), Inches(y), Inches(w), Inches(h))
    s.name = name
    s.rotation = rotation
    xfrm = s._element.spPr.xfrm
    if flip_h: xfrm.set('flipH', '1')
    if flip_v: xfrm.set('flipV', '1')
    if crop:
        s.crop_left, s.crop_top, s.crop_right, s.crop_bottom = crop
    return s


def build_assets():
    ASSETS.mkdir(parents=True, exist_ok=True)
    width, height = 2100, 1470
    y, x = np.indices((height, width))
    # Unique asymmetric orientation marks + near-Nyquist bands reveal whether
    # the restored image is source detail or an upsampled 200-dpi derivative.
    rgb = np.full((height, width, 3), 248, dtype=np.uint8)
    rgb[:height//2, :width//2] = (230, 245, 255)
    rgb[:height//2, width//2:] = (255, 235, 225)
    rgb[height//2:, :width//2] = (235, 255, 220)
    rgb[height//2:, width//2:] = (240, 225, 255)
    rgb[(x % 42 < 2) | (y % 42 < 2)] = (45, 45, 45)
    rgb[1120:1280, 100:1950] = np.where(((x[1120:1280, 100:1950]//2) % 2)[..., None], 255, 0)
    pattern = Image.fromarray(rgb)
    d = ImageDraw.Draw(pattern)
    font_path = Path('C:/Windows/Fonts/arial.ttf')
    font = ImageFont.truetype(str(font_path), 72) if font_path.exists() else ImageFont.load_default()
    small = ImageFont.truetype(str(font_path), 38) if font_path.exists() else ImageFont.load_default()
    for text, xy, fill in [('TL 1', (60, 40), 'blue'), ('TR 2', (1780, 40), 'red'),
                           ('BL 3', (60, 1320), 'green'), ('BR 4', (1780, 1320), 'purple')]:
        d.text(xy, text, fill=fill, font=font)
    d.text((540, 650), 'Original 2100 x 1470', font=font, fill='black')
    d.text((500, 920), '2-pixel frequency bars below', font=small, fill='black')
    d.line([(50, 400), (2050, 1000)], fill=(220, 0, 90), width=5)
    pattern.save(ASSETS/'pattern.png', dpi=(600,600))

    # Continuous alpha, fully transparent holes, and differently colored RGB
    # below transparent areas detect black/white fringes and incorrect masks.
    alpha = np.clip((x - 80) / 1840 * 255, 0, 255).astype(np.uint8)
    rgba = np.dstack((rgb, alpha))
    hole = (x-900)**2 + (y-740)**2 < 240**2
    rgba[hole, 3] = 0
    rgba[hole, :3] = (255, 0, 255)
    Image.fromarray(rgba, 'RGBA').save(ASSETS/'alpha.png', dpi=(600,600))

    icon = Image.new('RGBA', (52, 32), (0,0,0,0))
    d = ImageDraw.Draw(icon); d.rectangle((1,1,50,30), fill=(30,50,150,255), outline='black')
    d.line((2,28,49,2), fill='yellow', width=2)
    icon.save(ASSETS/'lowres.png')
    key = pattern.copy(); d=ImageDraw.Draw(key); d.rectangle((500,300,1600,1100), fill='#ff00ff')
    key.save(ASSETS/'colorkey.png',dpi=(600,600))


def build():
    build_assets()
    prs = Presentation()
    prs.slide_width = Inches(12); prs.slide_height = Inches(9)
    blank = prs.slide_layouts[6]
    case_records=[]

    def record(slide_no, shape, case, asset, expectation, **extra):
        im=Image.open(asset)
        case_records.append(dict(slide=slide_no,shape_id=shape.shape_id,case=case,
            asset=asset.name,pixels=list(im.size),mode=im.mode,
            source_sha256=hashlib.sha256(asset.read_bytes()).hexdigest(),
            expected=expectation,**extra))

    slide=prs.slides.add_slide(blank)
    label(slide.shapes,'Figure 1 (full width). Source pixels and layout',.2,.08,11.6,.35,19)
    for name,x in [('A plain',.3),('B cropped',4.2),('C rotated + flipped',8.1)]:
        label(slide.shapes,name,x,.52,3.5,.3,15)
    s=add_picture(slide.shapes,ASSETS/'pattern.png','HR plain',.3,.92,3.5,2.45)
    record(1,s,'plain',ASSETS/'pattern.png','exact original RGB pixels; no resampling')
    s=add_picture(slide.shapes,ASSETS/'pattern.png','HR crop',4.2,.92,3.5,2.45,crop=(.2,.2,.1,.1))
    record(1,s,'crop',ASSETS/'pattern.png','preserve visible source region and existing clip',crop=[.2,.2,.1,.1])
    s=add_picture(slide.shapes,ASSETS/'pattern.png','HR rotate90 flipH',8.85,.65,2.1,3.0,rotation=90,flip_h=True)
    record(1,s,'rotate90_flipH',ASSETS/'pattern.png','orientation marks TL/TR/BL/BR retain correct positions',rotation=90,flip_h=True)

    label(slide.shapes,'D alpha + arbitrary rotation',.3,3.75,3.6,.3,15)
    background(slide.shapes,.3,4.15,3.5,2.9,(35,145,185))
    s=add_picture(slide.shapes,ASSETS/'alpha.png','HR alpha rotated23',.7,4.55,2.7,1.89,rotation=23)
    record(1,s,'alpha_rotate23',ASSETS/'alpha.png','alpha holes/gradient show cyan background, no halo',rotation=23)

    label(slide.shapes,'E nested scaled group + foreground',4.2,3.75,3.9,.3,15)
    outer=slide.shapes.add_group_shape(); outer.name='Outer QA scaled group'
    inner=outer.shapes.add_group_shape(); inner.name='Inner QA group'
    s=add_picture(inner.shapes,ASSETS/'pattern.png','HR nested group',0,0,3.5,2.45)
    label(inner.shapes,'VECTOR FOREGROUND',.25,.95,3.0,.4,19,(200,0,40))
    bar=background(inner.shapes,.2,1.5,3.1,.04,(220,0,45))
    inner.left=Inches(.15); inner.top=Inches(.15)
    label(outer.shapes,'Nested group label',.1,2.7,3.4,.35,15)
    outer.left=Inches(4.35); outer.top=Inches(4.2); outer.width=Inches(3.4); outer.height=Inches(2.8)
    record(1,s,'nested_scaled_group',ASSETS/'pattern.png','native foreground and relative ordering remain unchanged',top_level_id=outer.shape_id,group_ids=[outer.shape_id,inner.shape_id])

    label(slide.shapes,'F low source pixels + alpha gradient',8.1,3.75,3.7,.3,15)
    background(slide.shapes,8.1,4.15,3.55,2.9,(255,208,65))
    s=add_picture(slide.shapes,ASSETS/'alpha.png','HR alpha unrotated',8.2,4.6,3.3,2.31)
    record(1,s,'alpha_unrotated',ASSETS/'alpha.png','exact source alpha retained over yellow background')
    s=add_picture(slide.shapes,ASSETS/'lowres.png','LOW native pixels',8.2,4.18,1.0,32/52)
    record(1,s,'low_resolution',ASSETS/'lowres.png','report below target; never upscale or invent detail')
    label(slide.shapes,'Only 52 x 32 px',9.35,4.2,2.2,.3,16)
    label(slide.shapes,'All colored labels, rectangles and lines remain native vector objects.',.3,7.45,11.3,.35,17)

    slide=prs.slides.add_slide(blank)
    label(slide.shapes,'Figure 2 (full width). Effects and clip safeguards',.2,.08,11.6,.35,19)
    for name,x in [('G color-key transparency',.3),('H brightness/contrast',4.2),('I ellipse mask + flipV',8.1)]:
        label(slide.shapes,name,x,.62,3.6,.35,15)
        background(slide.shapes,x,1.05,3.5,2.45,(95,180,135))
    s=add_picture(slide.shapes,ASSETS/'colorkey.png','EFFECT colorkey',.3,1.05,3.5,2.45)
    change=OxmlElement('a:clrChange'); change.set('useA','1')
    fr=OxmlElement('a:clrFrom'); srgb=OxmlElement('a:srgbClr'); srgb.set('val','FF00FF'); fr.append(srgb)
    to=OxmlElement('a:clrTo'); srgb=OxmlElement('a:srgbClr'); srgb.set('val','FF00FF'); al=OxmlElement('a:alpha'); al.set('val','0'); srgb.append(al); to.append(srgb)
    change.extend([fr,to]); s._element.blipFill.blip.append(change)
    record(2,s,'color_key',ASSETS/'colorkey.png','preserve transparency effect or retain native fallback with explicit unsupported warning')
    s=add_picture(slide.shapes,ASSETS/'pattern.png','EFFECT brightness',4.2,1.05,3.5,2.45)
    lum=OxmlElement('a:lum'); lum.set('bright','25000');lum.set('contrast','10000');s._element.blipFill.blip.append(lum)
    record(2,s,'brightness_contrast',ASSETS/'pattern.png','must not silently replace treated image with untreated RGB; faithful effect or explicit fallback')
    s=add_picture(slide.shapes,ASSETS/'pattern.png','HR ellipse clip flipV',8.1,1.05,3.5,2.45,flip_v=True)
    s._element.spPr.prstGeom.set('prst','ellipse')
    record(2,s,'ellipse_flipV',ASSETS/'pattern.png','ellipse clip and vertical flip retained; corners show green')
    label(slide.shapes,'J source alpha + 50% object opacity',.3,4.1,5.5,.35,15)
    background(slide.shapes,.3,4.6,5.0,3.5,(70,65,160))
    s=add_picture(slide.shapes,ASSETS/'alpha.png','EFFECT object opacity',.3,4.6,5.0,3.5)
    am=OxmlElement('a:alphaModFix'); am.set('amt','50000');s._element.blipFill.blip.append(am)
    record(2,s,'object_opacity',ASSETS/'alpha.png','source alpha times object opacity preserved or explicit fallback')
    label(slide.shapes,'Effects are intentional: an honest unresolved warning is safer than a visually changed image.',5.7,4.8,5.5,1.2,20)

    destination=ROOT/'highres_synthetic.pptx'
    prs.save(destination)
    expected={'source':destination.name,'slide_size_inches':[12,9],
        'purpose':'pixel and appearance preservation, not merely larger DPI numbers',
        'cases':case_records,'source_sha256':hashlib.sha256(destination.read_bytes()).hexdigest()}
    (ROOT/'highres_expected.json').write_text(json.dumps(expected,indent=2),encoding='utf-8')
    print(json.dumps({'pptx':str(destination),'cases':len(case_records),'slides':2},indent=2))


if __name__=='__main__':
    build()
