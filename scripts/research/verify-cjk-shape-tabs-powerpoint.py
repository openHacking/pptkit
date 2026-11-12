"""Verify first-open PowerPoint Chinese shape-tab source and PDF observations."""
import hashlib,json,plistlib,sys,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
import fitz
root=Path(__file__).resolve().parents[2];base=root/'.tmp/native-tab-stops/powerpoint-cjk-shape-GoFccp'
evidence=Path(next((x for x in sys.argv[1:] if x!='--record'),root/'docs/testing/tab-cjk-shape-powerpoint-evidence.json'))
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
source,pdf=base/'shapes.pptx',base/'shapes.pdf';assert sha(source)==json.loads((base/'manifest.json').read_text())['sha256']
a='{http://schemas.openxmlformats.org/drawingml/2006/main}';p='{http://schemas.openxmlformats.org/presentationml/2006/main}'
cases=[('left','l',['一','销售','总收入']),('center','ctr',['一','销售','总收入']),('right','r',['一','销售','总收入']),('decimal','dec',['1.2','123.45','0.678'])]
with zipfile.ZipFile(source) as z:
 assert z.testzip() is None
 for index,(kind,align,words) in enumerate(cases,1):
  slide=ET.fromstring(z.read(f'ppt/slides/slide{index}.xml'));shapes={s.find(p+'nvSpPr/'+p+'cNvPr').get('name'):s for s in slide.iter(p+'sp')}
  for i,word in enumerate(words):
   shape=shapes[f'{kind}-{i}'];para=shape.find(p+'txBody/'+a+'p');runs=para.findall(a+'r')
   assert [r.find(a+'t').text for r in runs]==['合计\t',word]
   assert all(r.find(a+'rPr/'+a+'latin').get('typeface')=='Hiragino Sans GB' and r.find(a+'rPr').get('sz')=='2800' for r in runs)
   assert para.find(a+'pPr/'+a+'tabLst/'+a+'tab').attrib=={'pos':'3556000','algn':align}
records=[]
with fitz.open(pdf) as doc:
 assert len(doc)==4
 for page,(kind,_,words) in zip(doc,cases):
  drawings=page.get_drawings();fills=[d for d in drawings if d['fill'] and all(abs(d['fill'][i]-v/255)<.01 for i,v in enumerate((242,245,250)))];assert len(fills)==3
  guides=[d for d in drawings if d['fill'] and d['fill'][0]>.7 and d['fill'][1]<.3];assert len(guides)==1;guide=(guides[0]['rect'].x0+guides[0]['rect'].x1)/2
  spans=[s for b in page.get_text('rawdict')['blocks'] if 'lines' in b for l in b['lines'] for s in l['spans']]
  dots=[c['origin'][0] for s in spans for c in s['chars'] if c['c']=='.'];rows=[]
  for i,word in enumerate(words):
   found=[s for s in spans if ''.join(c['c'] for c in s['chars'])==word];assert len(found)==1;s=found[0]
   assert s['font']==('HiraginoSansGB-W6' if i==2 else 'HiraginoSansGB-W3')
   chars=s['chars'];left,right=chars[0]['bbox'][0],chars[-1]['bbox'][2]
   anchor=left if kind=='left' else (left+right)/2 if kind=='center' else right if kind=='right' else dots[i];assert abs(anchor-guide)<.5
   xs=fills[i]['rect'].width/700;ys=fills[i]['rect'].height/65
   rows.append({'word':word,'bold':i==2,'font':s['font'],'anchor':anchor,'characters':[{'char':c['c'],'left':(c['bbox'][0]-fills[i]['rect'].x0)/xs,'right':(c['bbox'][2]-fills[i]['rect'].x0)/xs,'baseline':(c['origin'][1]-fills[i]['rect'].y0)/ys} for c in chars]})
  records.append({'kind':kind,'guide':guide,'rows':rows})
state=(base/'first.ax.txt').read_text();assert 'shapes.pptx' in state and 'Help: left-1, Value: 合计\t销售' in state
version=plistlib.loads(Path('/Applications/Microsoft PowerPoint.app/Contents/Info.plist').read_bytes())['CFBundleShortVersionString']
result={'scope':'First-open PowerPoint A4 PDF for 12 Chinese-prefix shape-tab samples, with Chinese or decimal suffix. Checks source declarations, selected PDF fonts and suffix characters. Only first page window captured; edit, save/reopen, Japanese/Korean and other fonts are not covered. Capture hashes do not independently prove provenance.', 'applicationVersion':version, 'sourceSha256':sha(source),'pdfSha256':sha(pdf),'axSha256':sha(base/'first.ax.txt'),'screenshotSha256':sha(base/'first.png'),'generatorSha256':sha(root/'scripts/prepare-native-cjk-shape-tabs.mjs'),'verifierSha256':sha(Path(__file__)),'pages':records}
if '--record' in sys.argv:evidence.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
else:assert json.loads(evidence.read_text())==result
print(json.dumps({'verified':True,'pages':4,'samples':12,'characters':sum(len(r['characters']) for p in records for r in p['rows'])}))
