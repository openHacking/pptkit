"""Per-script actual ZIP/PDF mutations for five native edit controls."""
import importlib.util, json, sys, tempfile, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz
spec=importlib.util.spec_from_file_location('multi',Path(__file__).with_name('real-font-multi.py'))
o=importlib.util.module_from_spec(spec);spec.loader.exec_module(o)
p=Path(sys.argv[1]);expected=json.loads((p/'expected.json').read_text());source=o.edit.read_package(p/'fresh-round3.pptx')
rejected=[]
def reject(name,fn):
 try:fn()
 except AssertionError:rejected.append(name);return
 raise AssertionError('Accepted corruption: '+name)
def rewrite(file,output,part,change):
 with zipfile.ZipFile(file) as z,zipfile.ZipFile(output,'w') as out:
  for entry in z.infolist():
   data=z.read(entry.filename)
   if entry.filename==part:
    root=ET.fromstring(data);body=root.findall('p:cSld/p:spTree/p:sp',o.edit.NS)[1].find('p:txBody',o.edit.NS)
    change(body);data=ET.tostring(root)
   out.writestr(entry,data)
def runs(body):return body.findall('a:p/a:r',o.edit.NS)
def style(body):return runs(body)[-1].find('a:rPr',o.edit.NS)
def merged(body):
 rs=runs(body);rs[0].find('a:t',o.edit.NS).text+=''.join(r.find('a:t',o.edit.NS).text for r in rs[1:])
 for r in rs[1:]:body.find('a:p',o.edit.NS).remove(r)
with tempfile.TemporaryDirectory(prefix='pptkit-multi-negative-') as temp:
 temp=Path(temp)
 for id in o.CASES:
  index=next(i for i,c in enumerate(source) if c['id']==id);part=source[index]['part']
  cases=[('text','edit',lambda b:setattr(runs(b)[-1].find('a:t',o.edit.NS),'text','Ψ')),
   ('language','edit',lambda b:style(b).set('lang','en-US')),
   ('restore-map','edit',lambda b:style(b).find('a:latin',o.edit.NS).set('typeface','Georgia')),
   ('merge-runs','edit',merged),('missing-bold','bold',lambda b:style(b).set('b','0')),
   ('missing-italic','italic',lambda b:style(b).set('i','0'))]
  for name,stage,change in cases:
   file=temp/(id+'-'+name+'.pptx');rewrite(p/('fresh-'+stage+'-multi.pptx'),file,part,change)
   reject(id+'/'+name,lambda:o.inspect_package(file,source,stage))
  doc=fitz.open(p/'fresh-edit-multi.pdf');doc[index].add_redact_annot(fitz.Rect(0,60,720,260),fill=(1,1,1));doc[index].apply_redactions()
  file=temp/(id+'-blank.pdf');doc.save(file);doc.close()
  reject(id+'/blank-body',lambda:o.inspect_pdf(file,expected,'edit'))
  doc=fitz.open(p/'fresh-italic-multi.pdf');changed=False
  for xref in doc[index].get_contents():
   stream=doc.xref_stream(xref)
   if b'67.96875' in stream:doc.update_stream(xref,stream.replace(b'67.96875',b'0'));changed=True
  assert changed;file=temp/(id+'-no-shear.pdf');doc.save(file);doc.close()
  reject(id+'/remove-shear',lambda:o.inspect_pdf(file,expected,'italic'))
 # A genuine old-theme package cannot satisfy the changed-theme stage.
 reject('unchanged-theme',lambda:o.inspect_package(p/'fresh-edit-multi.pptx',source,'theme'))
 reject('wrong-bold-face',lambda:o.inspect_pdf(p/'fresh-bold-multi.pdf',expected,'edit'))
 # Change a real PDF font resource name, not just expected metadata.
 doc=fitz.open(p/'fresh-edit-multi.pdf')
 xref=next(f[0] for f in doc[9].get_fonts() if f[3].endswith('+Cambria'))
 doc.xref_set_key(xref,'BaseFont','/Georgia');file=temp/'wrong-fallback.pdf';doc.save(file);doc.close()
 reject('wrong-fallback-face',lambda:o.inspect_pdf(file,expected,'edit'))
print(json.dumps({'rejectedCount':len(rejected),'rejectedMutations':rejected}))
