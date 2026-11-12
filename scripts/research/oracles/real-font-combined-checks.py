"""Actual same-paragraph source/PDF corruptions, including invisible-but-extractable text."""
import importlib.util,json,re,sys,tempfile,zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz
spec=importlib.util.spec_from_file_location('combined',Path(__file__).with_name('real-font-combined.py'))
o=importlib.util.module_from_spec(spec);spec.loader.exec_module(o)
p=Path(sys.argv[1]);e=json.loads((p/'expected.json').read_text());src=o.edit.read_package(p/'combined-fonts.pptx',1)[0]
initial=o.inspect_package(p/'combined-edit.pptx',e,src,'edit');rejected=[]
def reject(name,fn):
 try:fn()
 except AssertionError:rejected.append(name);return
 raise AssertionError('Accepted corruption: '+name)
def rewrite(stem,file,change):
 with zipfile.ZipFile(p/(stem+'.pptx')) as z,zipfile.ZipFile(file,'w') as out:
  for entry in z.infolist():
   data=z.read(entry.filename)
   if entry.filename==src['part']:
    root=ET.fromstring(data);b=root.findall('p:cSld/p:spTree/p:sp',o.edit.NS)[1].find('p:txBody',o.edit.NS);change(b);data=ET.tostring(root)
   out.writestr(entry,data)
def rs(b):return b.findall('a:p/a:r',o.edit.NS)
def st(b,i):return rs(b)[i].find('a:rPr',o.edit.NS)
def drop(b,i):b.find('a:p',o.edit.NS).remove(rs(b)[i])
def swap(b,a,c):
 parent=b.find('a:p',o.edit.NS);children=list(parent);i=children.index(rs(b)[a]);j=children.index(rs(b)[c]);children[i],children[j]=children[j],children[i];parent[:]=children
with tempfile.TemporaryDirectory(prefix='pptkit-combined-negative-') as temp:
 temp=Path(temp)
 for i,run in enumerate(e['expectedRuns']):
  for name,change in [('family',lambda b:st(b,i).find('a:latin',o.edit.NS).set('typeface','Wrong')),('language',lambda b:st(b,i).set('lang','wrong'))]:
   file=temp/(str(i)+name+'.pptx');rewrite('combined-edit',file,change)
   reject(run['fontId']+'/'+name,lambda:o.inspect_package(file,e,src,'edit'))
 for a,c in [(1,2),(2,3),(1,3)]:
  file=temp/('swap-'+str(a)+str(c)+'.pptx');rewrite('combined-edit',file,lambda b:swap(b,a,c));reject('source-run-order/'+str(a)+str(c),lambda:o.inspect_package(file,e,src,'edit'))
 for i in range(1,9):
  file=temp/('bleed-'+str(i)+'.pptx');rewrite('combined-bold',file,lambda b:st(b,i).set('b','1'));reject('local-style-bleed/'+str(i),lambda:o.inspect_package(file,e,src,'bold'))
 for stem,stage,name,change in [
  ('combined-bold','bold','missing-bold',lambda b:st(b,0).set('b','0')),
  ('combined-italic','italic','missing-italic',lambda b:st(b,0).set('i','0')),
  ('combined-edit','edit','missing-Greek',lambda b:drop(b,7)),
  ('combined-edit','edit','missing-X',lambda b:setattr(rs(b)[0].find('a:t',o.edit.NS),'text','Latin ')),
  ('combined-edit','edit','missing-Hebrew-marks',lambda b:setattr(rs(b)[6].find('a:t',o.edit.NS),'text','ש '))]:
  file=temp/(name+'.pptx');rewrite(stem,file,change);reject(name,lambda:o.inspect_package(file,e,src,stage))
 fonts={f['id']:f['postscriptName'] for f in e['fonts']}
 for run in e['expectedRuns']:
  font=fonts[run['fontId']]
  doc=fitz.open(p/'combined-edit.pdf');traces=[t for t in doc[0].get_texttrace() if t['font']==font and t['bbox'][1]>60 and any(c[0]!=32 for c in t['chars'])]
  assert traces
  for trace in traces:doc[0].draw_rect(fitz.Rect(trace['bbox']),color=None,fill=(1,1,1),overlay=True)
  file=temp/(run['fontId']+'-white-overlay.pdf');doc.save(file);doc.close()
  reject(run['fontId']+'/invisible-extractable-span',lambda:o.inspect_pdf(file,e,'edit'))
  doc=fitz.open(p/'combined-edit.pdf')
  for f in doc[0].get_fonts():
   if f[3].split('+')[-1]==font:doc.xref_set_key(f[0],'BaseFont','/Wrong')
  file=temp/(run['fontId']+'-resource.pdf');doc.save(file);doc.close()
  reject(run['fontId']+'/resource-name-conflict',lambda:o.inspect_pdf(file,e,'edit'))
 for a,b in [('sc','tc'),('tc','ja'),('sc','ja')]:
  doc=fitz.open(p/'combined-edit.pdf');resources={f[3].split('+')[-1]:f[4].encode() for f in doc[0].get_fonts()};left,right=resources[fonts[a]],resources[fonts[b]];mapping={left:right,right:left};count=0
  for xref in doc[0].get_contents():
   stream=doc.xref_stream(xref)
   def replace(match):
    global count
    if match[1] in mapping:count+=1;return b'/'+mapping[match[1]]+b' 1 Tf'
    return match[0]
   doc.update_stream(xref,re.sub(rb'/([A-Za-z0-9]+)\s+1\s+Tf',replace,stream))
  assert count>=2;file=temp/('actual-'+a+'-'+b+'.pdf');doc.save(file);doc.close()
  reject('actual-CJK-font-binding/'+a+b,lambda:o.inspect_pdf(file,e,'edit'))
 def check_round(file):o.assert_round_fields(o.inspect_package(file,e,src,'edit',1),initial,'edit')
 file=temp/'proofing-flag.pptx';rewrite('combined-edit-round1',file,lambda b:st(b,0).attrib.pop('err'));reject('unexpected-proofing-fields',lambda:check_round(file))
 def merge_space(b):
  rs(b)[0].find('a:t',o.edit.NS).text+=' ';drop(b,1)
 file=temp/'proofing-topology.pptx';rewrite('combined-edit-round1',file,merge_space);reject('unexpected-proofing-topology',lambda:check_round(file))
 reject('unchanged-theme',lambda:o.inspect_package(p/'combined-edit.pptx',e,src,'theme'))
 reject('actual-italic-as-bold',lambda:o.inspect_pdf(p/'combined-italic.pdf',e,'bold'))
print(json.dumps({'rejectedCount':len(rejected),'rejectedMutations':rejected}))
