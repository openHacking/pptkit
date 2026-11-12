"""Generate an independent PowerPoint XY scatter research deck.

The python-pptx generator creates the native chart and XLSX. Only the
chart-wide blank-display scalar is changed after creation. This is a research
fixture, not PPTKit scatter output.
"""
import hashlib, json, os, tempfile
from pathlib import Path
from pptx import Presentation
from pptx.chart.data import XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_MARKER_STYLE
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

root = Path(__file__).resolve().parents[1]
base = root / '.tmp/native-scatter-missing-x'
base.mkdir(parents=True, exist_ok=True)
dirpath = Path(tempfile.mkdtemp(prefix='powerpoint-', dir=base))
prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333333), Inches(7.5)
examples = [
    ('markers-gap', XL_CHART_TYPE.XY_SCATTER, 'gap'),
    ('line-gap', XL_CHART_TYPE.XY_SCATTER_LINES, 'gap'),
    ('line-zero', XL_CHART_TYPE.XY_SCATTER_LINES, 'zero'),
    ('line-span', XL_CHART_TYPE.XY_SCATTER_LINES, 'span'),
]
series = [('RED', [(-8,4),(None,3),(2,-4),(8,6)], RGBColor(255,0,0)),
          ('BLUE',[(7,-6),(-1,4),(3,2),(-7,-2)], RGBColor(0,0,255))]
for name, chart_type, mode in examples:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title = slide.shapes.add_textbox(Inches(.4), Inches(.1), Inches(6), Inches(.4))
    title.name = 'title-' + name
    title.text_frame.text = name
    title.text_frame.paragraphs[0].runs[0].font.size = Pt(16)
    data = XyChartData()
    for label, pairs, _ in series:
        s = data.add_series(label)
        for x, y in pairs:
            s.add_data_point(x, y)
    frame = slide.shapes.add_chart(chart_type, Inches(1), Inches(.75), Inches(10), Inches(5.5), data)
    frame.name = 'chart-' + name
    chart = frame.chart
    chart.has_legend = False
    for axis in (chart.category_axis, chart.value_axis):
        axis.minimum_scale, axis.maximum_scale, axis.major_unit = -10, 10, 5
        axis.has_major_gridlines = False
    for index, (_, _, color) in enumerate(series):
        s = chart.series[index]
        s.marker.style = XL_MARKER_STYLE.CIRCLE
        s.marker.size = 10
        s.marker.format.fill.solid()
        s.marker.format.fill.fore_color.rgb = color
        s.marker.format.line.fill.background()
        if chart_type == XL_CHART_TYPE.XY_SCATTER:
            s.format.line.fill.background()
        else:
            s.format.line.color.rgb = color
            s.format.line.width = Pt(2)
    node = chart._chartSpace.chart.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}dispBlanksAs')
    if node is None:
        node = OxmlElement('c:dispBlanksAs')
        chart._chartSpace.chart.append(node)
    node.set('val', mode)
filename = dirpath / 'scatter-missing-x.pptx'
prs.save(str(filename))
manifest = {'generator':'python-pptx', 'scope':'independent XY fixture; blank-display scalar set explicitly', 'groups':[x[0] for x in examples], 'sha256':hashlib.sha256(filename.read_bytes()).hexdigest(), 'series':[[label, pairs] for label,pairs,_ in series]}
(dirpath/'manifest.json').write_text(json.dumps(manifest, indent=2)+'\n')
print(json.dumps({'directory':str(dirpath),'file':str(filename),'sha256':manifest['sha256']}))
