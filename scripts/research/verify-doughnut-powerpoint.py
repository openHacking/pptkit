"""Check retained PowerPoint doughnut sample, source/saved payloads, captures and first print."""
import hashlib, json, plistlib, sys, zipfile, xml.etree.ElementTree as ET
from pathlib import Path
import fitz
root = Path(__file__).resolve().parents[2]
base = root / '.tmp/native-doughnut/powerpoint-TFHkIb'
evidence = Path(next((v for v in sys.argv[1:] if v != '--record'), root / 'docs/testing/chart-doughnut-powerpoint-evidence.json'))
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
c = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
source, saved = base/'doughnut.pptx', base/'doughnut-round-1.pptx'
assert sha(source) == json.loads((base/'manifest.json').read_text())['sha256']
packages = []
for file in (source, saved):
    with zipfile.ZipFile(file) as z:
        assert z.testzip() is None
        charts = []
        for i, (hole, angle, values) in enumerate([(50, 0, [[1,3],[3,1]]), (30,90,[[1,3]]), (70,0,[[2,None]])], 1):
            xml = ET.fromstring(z.read(f'ppt/charts/chart{i}.xml'))
            chart = xml.find('.//'+c+'doughnutChart'); assert chart is not None
            assert int(chart.find(c+'holeSize').get('val')) == hole and int(chart.find(c+'firstSliceAng').get('val')) == angle
            series = chart.findall(c+'ser'); assert len(series) == len(values)
            for ser, expected in zip(series, values):
                cache = ser.find('.//'+c+'val/'+c+'numRef/'+c+'numCache'); assert cache.find(c+'ptCount').get('val') == '2'
                points = {int(v.get('idx')): float(v.find(c+'v').text) for v in cache.findall(c+'pt')}
                assert points == {j: v for j, v in enumerate(expected) if v is not None}
            charts.append({'holeSize':hole,'firstSliceAngle':angle,'seriesValues':values})
        packages.append({'sha256':sha(file), 'charts':charts})
with fitz.open(base/'doughnut.pdf') as doc:
    assert len(doc) == 3
    def colored(page):
        return {tuple(round(x, 3) for x in d['fill']): d['rect'] for d in page.get_drawings() if d.get('fill') and max(d['fill']) > .5 and min(d['fill']) < .1}
    a,b,d = [colored(page) for page in doc]
    red, blue, green, orange = (1.,0.,0.), (0.,0.,1.), (0.,.667,0.), (1.,.667,0.)
    assert set(a) == {red, blue, green, orange} and set(b) == {red, blue} and set(d) == {red}
    outer = a[green]; inner = a[blue]; quarter = a[red]; cx,cy = (outer.x0+outer.x1)/2,(outer.y0+outer.y1)/2
    # Inner radius 75% of outer; red occupies upper-right quarter, orange upper-left outer quarter.
    assert abs(inner.width/outer.width - .75) < .01 and abs(inner.height/outer.height - .75) < .01
    assert abs(quarter.x0-cx) < .1 and abs(quarter.y1-cy) < .1 and abs(a[orange].x1-cx) < .1 and abs(a[orange].y1-cy) < .1
    # 90-degree first slice occupies lower-right quarter.
    rotated = b[red]; whole = b[blue]; rx,ry = (whole.x0+whole.x1)/2,(whole.y0+whole.y1)/2
    assert abs(rotated.x0-rx) < .1 and abs(rotated.y0-ry) < .1
    # Slide background is the first white rectangle. Transform the native paths back
    # to the authored 960x540 slide and compare against the independently computed
    # plot center and radius (700x420 chart with 8pt inset on all sides).
    comparisons=[]
    for page, draws, outer_color, hole in zip(doc, (a,b,d), (green,blue,red), (50,30,70)):
        frame = next(x['rect'] for x in page.get_drawings() if x.get('fill') == (1.,1.,1.) and x['rect'].width > 700)
        sx,sy=frame.width/960,frame.height/540; box=draws[outer_color]
        x,y=((box.x0+box.x1)/2-frame.x0)/sx,((box.y0+box.y1)/2-frame.y0)/sy
        rx,ry=box.width/(2*sx),box.height/(2*sy)
        assert max(abs(x-430),abs(y-280),abs(rx-202),abs(ry-202)) < .3
        comparisons.append({'nativeCenter':[x,y], 'nativeRadius':[rx,ry], 'previewCenter':[430,280], 'previewRadius':202, 'holeSize':hole})
    measurements = {'outerBox':list(outer), 'innerBox':list(inner), 'innerToOuterRadius':inner.width/outer.width, 'rotatedRedBox':list(rotated), 'blankColors':[list(x) for x in d], 'nativeVsPreview':comparisons}
captures=[]
for phase in ('first','reopen'):
    for i,name in enumerate(('rings','rotate','blank'),1):
        ax=base/f'{phase}-{i}.ax.txt'; text=ax.read_text(); assert f'chart-{name}' in text and f'第 {i} 张幻灯片' in text
        captures.append({'phase':phase,'slide':i,'axSha256':sha(ax),'screenshotSha256':sha(base/f'{phase}-{i}.png')})
version=plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString']
result={'scope':'Three generated charts observed on first open and one save/close/reopen in PowerPoint. First-open PDF checks colored ring order, quarter orientation and blank plot. Package inspection checks geometry and cached values; screenshots/AX hashes alone do not prove provenance. No native data edits, reopened print, alternate producers or general layout parity are established.', 'applicationVersion':version,'generatorSha256':sha(root/'scripts/prepare-native-doughnut.mjs'),'verifierSha256':sha(Path(__file__)),'packages':packages,'firstPrintSha256':sha(base/'doughnut.pdf'),'measurements':measurements,'captures':captures}
if '--record' in sys.argv: evidence.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
else: assert json.loads(evidence.read_text()) == result
print(json.dumps({'verified':True,'charts':3,'reopenRounds':1,'firstPrintPages':3}))
