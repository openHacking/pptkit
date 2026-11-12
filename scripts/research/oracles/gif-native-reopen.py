#!/usr/bin/env python3
"""Verify captured artifacts, not PowerPoint UI provenance or precise timing."""
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import json
import posixpath
import xml.etree.ElementTree as ET
from PIL import Image

ROOT = Path(__file__).resolve().parents[3]
BASE = ROOT / '.tmp/run-_2s_4erk'
NS = {'p':'http://schemas.openxmlformats.org/presentationml/2006/main', 'a':'http://schemas.openxmlformats.org/drawingml/2006/main', 'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'rel':'http://schemas.openxmlformats.org/package/2006/relationships'}
def check(condition, message):
    if not condition: raise ValueError(message)
def digest(b): return sha256(b).hexdigest()
def package(data, gif):
    with ZipFile(BytesIO(data)) as z:
        check(len(z.namelist()) == len(set(z.namelist())), 'duplicate parts')
        slide = ET.fromstring(z.read('ppt/slides/slide1.xml'))
        rels = ET.fromstring(z.read('ppt/slides/_rels/slide1.xml.rels'))
        image_rels = {r.get('Id'): r.get('Target') for r in rels.findall('rel:Relationship', NS)
                      if r.get('Type','').endswith('/image') and r.get('TargetMode') is None}
        pictures = slide.findall('.//p:pic', NS); check(len(pictures) == 2, 'picture count')
        result = []
        for pic in pictures:
            nv = pic.find('p:nvPicPr/p:cNvPr', NS); blip = pic.find('p:blipFill/a:blip', NS)
            ref = blip.get('{%s}embed' % NS['r']); check(ref in image_rels, 'missing internal image')
            target = posixpath.normpath(posixpath.join('ppt/slides', image_rels[ref]))
            check(target.startswith('ppt/media/') and z.read(target) == gif, 'GIF bytes')
            xfrm = pic.find('p:spPr/a:xfrm', NS); off=xfrm.find('a:off', NS); ext=xfrm.find('a:ext', NS)
            result.append({'id': nv.get('id'), 'name': nv.get('name'), 'description': nv.get('descr'), 'target': target, 'relationship': ref,
                           'box': [int(off.get('x')), int(off.get('y')), int(ext.get('cx')), int(ext.get('cy'))]})
        check(result[0]['target'] == result[1]['target'], 'lost shared media')
        check([(r['id'],r['name'],r['description'],r['box']) for r in result] == [
            ('3','gif-left','gif-left',[1143000,1905000,3048000,3048000]),
            ('4','gif-right','gif-right',[6604000,1905000,3048000,3048000])], 'identity or geometry')
        check('Animated GIF — native playback' in ''.join(t.text or '' for t in slide.findall('.//a:t',NS)), 'title')
        return result

def captures(rows, shots):
    check(len(rows) == len(shots) == 90, 'capture count')
    changes=[]; last=None
    for i,(row,data) in enumerate(zip(rows,shots)):
        check(row['file'] == f'reopen-{i}.png' and row['start'] <= row['end'], 'capture metadata')
        if i: check(row['start'] >= rows[i-1]['end'], 'capture order')
        im=Image.open(BytesIO(data)).convert('RGB'); check(im.size == (1920,1080), 'size')
        def classify(pixel):
            r,g,b=pixel
            return 'red' if r>200 and g<80 and b<80 else 'blue' if b>200 and r<80 and g<80 else 'other'
        colors=[classify(im.getpixel(point)) for point in [(400,500),(1200,500)]]
        check(colors[0] == colors[1] and colors[0] != 'other', 'sample colors')
        if colors[0] != last: changes.append({'file':row['file'],'color':colors[0]}); last=colors[0]
    check(len(changes)>=3 and changes[0]['color']==changes[2]['color'] and changes[0]['color']!=changes[1]['color'], 'no observed return')
    return changes

def alter(pptx, name, change):
    out=BytesIO()
    with ZipFile(BytesIO(pptx)) as source, ZipFile(out,'w',ZIP_DEFLATED) as target:
        for item in source.infolist():
            data=source.read(item.filename); target.writestr(item.filename, change(data) if item.filename==name else data)
    return out.getvalue()

def main():
    prior=json.loads((ROOT/'docs/testing/gif-powerpoint-playback-evidence.json').read_text()); gif=(BASE/'motion.gif').read_bytes()
    check(digest(gif)==prior['gifSha256'], 'source hash')
    files=['candidate.pptx','native-round1.pptx','native-round2.pptx','native-round3.pptx']
    data=[(BASE/file).read_bytes() for file in files]; check(digest(data[0]) == prior['pptxSha256'], 'candidate hash')
    observations=[package(b,gif) for b in data]
    rows=json.loads((BASE/'reopen-times.json').read_text()); shots=[(BASE/row['file']).read_bytes() for row in rows]; transitions=captures(rows,shots)
    failed=0
    for corrupt in [alter(data[-1],'ppt/media/image1.gif',lambda b:b[:-1]),
                    alter(data[-1],'ppt/slides/slide1.xml',lambda b:b.replace(b'gif-right',b'gif-wrong'))]:
        try: package(corrupt,gif)
        except Exception: failed+=1
    for bad in [shots[:-1], [shots[0]] * len(shots)]:
        try: captures(rows,bad)
        except Exception: failed+=1
    check(failed==4,'negative checks')
    evidence={'scope':'PowerPoint 16.104 three distinct Save As/close/reopen UI observations; file and pixel oracle cannot witness UI actions or establish exact timings.',
              'directory':str(BASE.relative_to(ROOT)), 'sourceGifSha256':digest(gif),
              'packages':[{'file':name,'sha256':digest(b),'pictures':obs} for name,b,obs in zip(files,data,observations)],
              'captureTimesSha256':digest((BASE/'reopen-times.json').read_bytes()),
              'screenshots':[{'file':r['file'],'sha256':digest(b)} for r,b in zip(rows,shots)],
              'transitions':transitions, 'negativeCasesRejected':failed,'oracleSha256':digest(Path(__file__).read_bytes())}
    (ROOT/'docs/testing/gif-powerpoint-reopen-evidence.json').write_text(json.dumps(evidence,indent=2)+'\n')
    print(json.dumps({'packages':len(data),'transitions':transitions,'negativeCasesRejected':failed}))
if __name__=='__main__': main()
