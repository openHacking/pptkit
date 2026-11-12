"""Real per-page XML/PDF corruptions plus retained native selection counterexamples."""
import importlib.util,json,sys,tempfile,zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz
spec=importlib.util.spec_from_file_location('mixed',Path(__file__).with_name('real-font-mixed.py'))
o=importlib.util.module_from_spec(spec);spec.loader.exec_module(o)
p=Path(sys.argv[1]);expected=json.loads((p/'expected.json').read_text());source=o.edit.read_package(p/'mixed-fonts.pptx',6)
rejected=[]
def reject(name,fn):
 try:fn()
 except AssertionError:rejected.append(name);return
 raise AssertionError('Accepted artifact corruption: '+name)
def rewrite(file,output,part,change):
 with zipfile.ZipFile(file) as z,zipfile.ZipFile(output,'w') as out:
  for entry in z.infolist():
   data=z.read(entry.filename)
   if entry.filename==part:
    root=ET.fromstring(data);body=root.findall('p:cSld/p:spTree/p:sp',o.edit.NS)[1].find('p:txBody',o.edit.NS);change(body);data=ET.tostring(root)
   out.writestr(entry,data)
def runs(b):return b.findall('a:p/a:r',o.edit.NS)
def style(b,i):return runs(b)[i].find('a:rPr',o.edit.NS)
def remove(b,i):b.find('a:p',o.edit.NS).remove(runs(b)[i])
with tempfile.TemporaryDirectory(prefix='pptkit-mixed-negative-') as temp:
 temp=Path(temp)
 for index,control in enumerate(expected['controls']):
  id=control['id'];part=source[index]['part']
  cases=[('missing-X','edit',lambda b:remove(b,1)),('new-language','edit',lambda b:style(b,1).set('lang','en-US')),
   ('wrong-Latin-family','edit',lambda b:style(b,1).find('a:latin',o.edit.NS).set('typeface','Georgia')),
   ('missing-Greek','edit',lambda b:remove(b,4)),
   ('native-text','edit',lambda b:setattr(runs(b)[3].find('a:t',o.edit.NS),'text','Wrong')),
   ('missing-local-bold','bold',lambda b:style(b,1).set('b','0')),
   ('bold-bleed-to-native','bold',lambda b:style(b,3).set('b','1')),
   ('italic-bleed-to-Greek','italic',lambda b:style(b,4).set('i','1'))]
  for name,stage,change in cases:
   file=temp/(id+'-'+name+'.pptx');rewrite(p/(o.STAGES[stage]+'.pptx'),file,part,change)
   reject(id+'/'+name,lambda:o.inspect_package(file,expected,source,stage))
  doc=fitz.open(p/'mixed-edit.pdf')  # The selected page owns its actual font resource.
  native_name=next(f['postscriptName'] for f in expected['fonts'] if f['id']==control['fontId'])
  xref=next(f[0] for f in doc[index].get_fonts() if f[3].split('+')[-1]==native_name)
  doc.xref_set_key(xref,'BaseFont','/Georgia');file=temp/(id+'-wrong-actual-native.pdf');doc.save(file);doc.close()
  reject(id+'/resource-face-conflict',lambda:o.inspect_pdf(file,expected,'edit'))
  doc=fitz.open(p/'mixed-edit.pdf')
  for font in doc[index].get_fonts():
   if font[3].split('+')[-1]==native_name:doc.xref_set_key(font[0],'BaseFont','/Georgia')
  file=temp/(id+'-all-native-resources.pdf');doc.save(file);doc.close()
  reject(id+'/all-native-faces-wrong',lambda:o.inspect_pdf(file,expected,'edit'))
  doc=fitz.open(p/'mixed-edit.pdf');traces=[t for t in doc[index].get_texttrace() if t['font']==native_name and t['bbox'][1]>60]
  assert traces
  for trace in traces:doc[index].add_redact_annot(fitz.Rect(trace['bbox']),fill=(1,1,1))
  doc[index].apply_redactions();file=temp/(id+'-missing-native.pdf');doc.save(file);doc.close()
  reject(id+'/missing-native-glyphs',lambda:o.inspect_pdf(file,expected,'edit'))
 for stem in ['mixed-bold','mixed-bold-range']:
  reject(stem+'/native-source-counterexample',lambda:o.inspect_package(p/(stem+'.pptx'),expected,source,'bold'))
  reject(stem+'/native-PDF-counterexample',lambda:o.inspect_pdf(p/(stem+'.pdf'),expected,'bold'))
 reject('unchanged-theme',lambda:o.inspect_package(p/'mixed-edit.pptx',expected,source,'theme'))
 reject('actual-italic-as-bold',lambda:o.inspect_pdf(p/'mixed-italic.pdf',expected,'bold'))
print(json.dumps({'rejectedCount':len(rejected),'rejectedMutations':rejected}))
