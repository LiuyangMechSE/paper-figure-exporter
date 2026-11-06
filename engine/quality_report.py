"""Human-readable raster audit; no image modification or resolution claims."""
import csv
import html
import json
from pathlib import Path


def write_quality_report(out, manifest):
    out = Path(out)
    esc = lambda value: html.escape(str(value))
    fmt = lambda value: '—' if value is None else f'{value:.1f}'
    entries = [c for c in manifest['candidates'] if c.get('image_quality')]
    target = manifest.get('quality_settings', {}).get('target_ppi', 600)
    data = {'source_sha256':manifest['source_sha256'], 'settings':manifest.get('quality_settings'),
            'figures':[{'output':c.get('output'), 'slide':c['slide'], 'width':c.get('width'),
                        'width_mm':c.get('output_width_mm'), 'preservation':c.get('image_preservation'),
                        'quality':c['image_quality']} for c in entries]}
    (out / 'quality_report.json').write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    keys = ['figure','slide','image','width_px','height_px','width_mm','height_mm','ppi_x','ppi_y',
            'effective_min_ppi','small_element','meets_target','required_width_px','required_height_px',
            'max_scale_at_target','soft_mask','potentially_visible']
    with (out / 'quality_report.csv').open('w', newline='', encoding='utf-8-sig') as stream:
        writer = csv.DictWriter(stream, fieldnames=keys); writer.writeheader()
        for c in entries:
            for im in c['image_quality']['images']:
                writer.writerow(dict(zip(keys, [c.get('output'),c['slide'],im.get('name',im.get('location')),
                    im.get('image_width_px'),im.get('image_height_px'),im.get('width_mm'),im.get('height_mm'),
                    im.get('dpi_x'),im.get('dpi_y'),im.get('min_ppi'),im.get('small_element'),im.get('meets_target'),
                    im.get('required_width_px'),im.get('required_height_px'),im.get('max_scale_at_target'),
                    json.dumps(im.get('masks'),ensure_ascii=False),im.get('potentially_visible')])))
    body = ['<!doctype html><meta charset="utf-8"><title>Figure quality report</title>',
        '<style>body{font:16px system-ui;max-width:1200px;margin:36px auto;padding:0 20px;color:#223}table{border-collapse:collapse;width:100%}td,th{padding:9px;text-align:left;border-bottom:1px solid #ddd}pre{white-space:pre-wrap;overflow-wrap:anywhere}details{margin:16px 0}.warning{color:#8a4b00}</style>',
        '<h1>图件清晰度报告 / Figure quality</h1>',
        f'<p>目标：最终物理尺寸下图片至少 {target:g} ppi。PDF 本身没有统一 DPI；矢量线条和文字没有固定像素分辨率。</p>',
        '<p>本工具保留原始图片像素，不插值放大、不自动缩小图宽来制造达标结果。PNG 预览的渲染 DPI 与 PDF 内图片的实际 PPI 是两个指标。</p>',
        '<p>“主体最低”排除报告中标记的小面积元素；“全部最低”包括小图标和透明蒙版。遮挡和非矩形裁切的可见范围采用保守估计。</p>',
        '<p><a href="review.html">查看分组与版式预览</a> · <a href="quality_report.csv">逐图片 CSV</a> · <a href="quality_report.json">完整数据 JSON</a></p>',
        '<table><tr><th>图件</th><th>宽度 mm</th><th>全部最低 ppi</th><th>主体最低 ppi</th><th>低于目标</th><th>判定</th></tr>']
    for c in entries:
        q=c['image_quality']
        status = '达标' if q['meets_target'] and q['assessment_complete'] else '需要检查 / 源像素不足'
        body.append(f'<tr><td><a href="{esc(c["output"])}">{esc(c["output"])}</a></td><td>{fmt(c.get("output_width_mm"))}</td><td>{fmt(q.get("min_ppi"))}</td><td>{fmt(q.get("min_large_ppi"))}</td><td>{len(q["below_target_images"])}/{q["image_count"]}</td><td>{status}</td></tr>')
    body.append('</table><h2>如何真正提高分辨率</h2><p>优先从 Origin / MATLAB / 绘图软件重新导出矢量 PDF、SVG 或 EMF，并保留可编辑文字；照片和位图使用更高像素的原始文件。若版式允许，可以减小最终图宽。下列尺寸是按现有可用像素计算的上限，不会自动应用。</p>')
    for c in entries:
        q=c['image_quality']
        body.append(f'<h2>{esc(c["output"])} · Slide {c["slide"]}</h2>')
        if q['below_target_images']:
            body.append('<table><tr><th>图片</th><th>现有像素</th><th>当前最低 ppi</th><th>当前尺寸所需像素</th><th>整图允许最大宽度 mm</th></tr>')
            for im in q['below_target_images']:
                scale=im.get('max_scale_at_target') if im.get('mask_resolution_known') else None
                maximum=c.get('output_width_mm',0)*scale if scale is not None else None
                required=f'{im.get("required_width_px")} × {im.get("required_height_px")}'
                for mask in im.get('masks',[]):
                    required+=f'<br>Mask {esc(mask.get("kind"))}: {mask.get("required_width_px")} × {mask.get("required_height_px")}'
                identity=str(im.get('location',''))+'/'+str(im.get('name',''))
                body.append(f'<tr><td>{esc(identity)}</td><td>{im.get("image_width_px")} × {im.get("image_height_px")}</td><td>{fmt(im.get("min_ppi"))}</td><td>{required}</td><td>{fmt(maximum)}</td></tr>')
            body.append('</table>')
        for issue in q.get('issues',[]):
            body.append('<p class="warning">'+esc(issue)+'</p>')
        body.append('<details><summary>原始像素保留 / 不支持对象及原因</summary><pre>'+esc(json.dumps(c.get('image_preservation'),ensure_ascii=False,indent=2))+'</pre></details>')
        body.append('<details><summary>逐图片及矢量、字体测量</summary><pre>'+esc(json.dumps(q,ensure_ascii=False,indent=2))+'</pre></details>')
    (out / 'quality_report.html').write_text('\n'.join(body),encoding='utf-8')
