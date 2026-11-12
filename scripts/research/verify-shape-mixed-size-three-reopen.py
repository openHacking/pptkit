"""Verify three PowerPoint save/reopen rounds and prints for mixed-size shape tab samples."""
import hashlib
import json
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
import fitz
root=Path(__file__).resolve().parents[2]
base=root/'.tmp/native-tab-stops/powerpoint-shape-mixed-size-774YI9'
evidence=Path(next((arg for arg in sys.argv[1:] if arg!='--record'),root/'docs/testing/tab-shape-mixed-size-three-reopen-powerpoint-evidence.json'))
sha=lambda path:hashlib.sha256(path.read_bytes()).hexdigest()
prior=root/'docs/testing/tab-shape-mixed-size-powerpoint-evidence.json'; first=json.loads(prior.read_text())
source,pdf=base/'shapes.pptx',base/'shapes.pdf'; saved=[base/f'shapes-round-{i}.pptx' for i in range(1,4)]; printed=[base/f'shapes-round-{i}.pdf' for i in (1,3)]
assert sha(source)==first['sourceSha256'] and sha(pdf)==first['pdfSha256']
a='{http://schemas.openxmlformats.org/drawingml/2006/main}';p='{http://schemas.openxmlformats.org/presentationml/2006/main}'
cases=[('left','l',['I','WIDE','MIXED']),('center','ctr',['I','WIDE','MIXED']),('right','r',['I','WIDE','MIXED']),('decimal','dec',['1.2','123.45','0.678'])]
for path in saved:
 with zipfile.ZipFile(path) as package:
  assert package.testzip() is None
  for index,(kind,align,words) in enumerate(cases,1):
   slide=ET.fromstring(package.read(f'ppt/slides/slide{index}.xml'))
   shapes={s.find(p+'nvSpPr/'+p+'cNvPr').get('name'):s for s in slide.iter(p+'sp')}
   for i,word in enumerate(words):
    shape=shapes[f'{kind}-{i}']; para=shape.find(p+'txBody/'+a+'p'); runs=para.findall(a+'r')
    assert shape.find(p+'nvSpPr/'+p+'cNvSpPr').get('txBox') is None
    assert shape.find(p+'spPr/'+a+'solidFill/'+a+'srgbClr').get('val')=='F2F5FA'
    assert [r.find(a+'t').text for r in runs]==['P\t',word]
    assert [r.find(a+'rPr').get('sz') for r in runs]==['2000',str([2800,3600,1600][i])]
    assert para.find(a+'pPr/'+a+'tabLst/'+a+'tab').attrib=={'pos':'3556000','algn':align}
for printed_file in printed:
 with fitz.open(pdf) as original, fitz.open(printed_file) as reopened:
  assert len(original)==len(reopened)==4
  for first_page,later in zip(original,reopened):
   assert tuple(first_page.rect)==tuple(later.rect) and first_page.get_text('words')==later.get_text('words')
   def fills(page): return [(tuple(d['rect']),d['fill']) for d in page.get_drawings() if d['fill'] and all(abs(d['fill'][i]-v/255)<.01 for i,v in enumerate((242,245,250)))]
   assert len(fills(first_page))==3 and fills(first_page)==fills(later)
rounds=[]
for i,path in enumerate(saved,1):
 ax=base/('reopen.ax.txt' if i==1 else f'reopen-{i}.ax.txt'); screenshot=base/('reopen.png' if i==1 else f'reopen-{i}.png')
 state=ax.read_text(); assert f'shapes-round-{i}.pptx' in state and 'Help: left-1, Value: P\tWIDE' in state
 rounds.append({'pptxSha256':sha(path),'axSha256':sha(ax),'screenshotSha256':sha(screenshot)})
result={'scope':'Three sequential PowerPoint save/close/reopen rounds of 12 mixed-size shape-text samples. Every saved package is checked; first and third reopened four-page prints match first-open word geometry and fills. Only first-page windows observed; no native edit or other-font/transform claim. Capture hashes do not independently establish provenance.', 'firstEvidenceSha256':sha(prior), 'rounds':rounds, 'reopenedPdfSha256':[sha(path) for path in printed], 'verifierSha256':sha(Path(__file__))}
if '--record' in sys.argv:evidence.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
else:assert json.loads(evidence.read_text())==result
print(json.dumps({'verified':True,'pages':4,'samples':12,'reopenRounds':3}))
