"""Actual mutations of source selector chains and native PDF observations."""
import importlib.util,json,sys,tempfile,zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz
spec=importlib.util.spec_from_file_location('table',Path(__file__).with_name('native-table-fonts.py'))
o=importlib.util.module_from_spec(spec);spec.loader.exec_module(o)
p=Path(sys.argv[1]);e=json.loads((p/'expected.json').read_text());source=o.inspect_package(p/'table-fonts.pptx',e,True);rejected=[]
selected={'minor-latin','minor-tc','minor-ko','minor-ko-kr','major-tc','major-ko','major-ko-kr','collection-tc','collection-ar','collection-generic-tc','collection-generic-ar'}
def reject(name,fn):
 try:fn()
 except AssertionError:rejected.append(name);return
 raise AssertionError('Accepted corruption: '+name)
def rewrite(file,part,change,stem='table-fonts'):
 with zipfile.ZipFile(p/(stem+'.pptx')) as z,zipfile.ZipFile(file,'w') as out:
  for entry in z.infolist():
   data=z.read(entry.filename)
   if entry.filename==part:
    root=ET.fromstring(data);change(root);data=ET.tostring(root)
   out.writestr(entry,data)
def table(root):return root.find('.//a:tbl',o.NS)
def run(root):return table(root).find('.//a:r',o.NS)
def normalized(file):
 actual=o.inspect_package(file,e);assert actual['controls']==o.expected_native_records(source['controls'])
with tempfile.TemporaryDirectory(prefix='pptkit-table-font-negative-') as temp:
 temp=Path(temp)
 for index,c in enumerate(e['controls']):
  if c['id'] not in selected:continue
  row=source['controls'][index];part=row['part'];id=c['id']
  actions=[('text',lambda r:setattr(run(r).find('a:t',o.NS),'text','Wrong')),
   ('language',lambda r:run(r).find('a:rPr',o.NS).set('lang','ko-KR' if c['language']=='ko' else 'ko' if c['language']=='ko-KR' else 'wrong')),
   ('style-binding',lambda r:setattr(table(r).find('a:tblPr/a:tableStyleId',o.NS),'text','{00000000-0000-0000-0000-000000000000}'))]
  for name,change in actions:
   file=temp/(id+'-'+name+'.pptx');rewrite(file,part,change);reject(id+'/'+name,lambda:o.inspect_package(file,e))
  def rule(root):
   if row['selector']['type']=='reference':node=root.find('a:themeElements/a:fontScheme/a:'+c['kind']+'Font',o.NS)
   else:node=next(s for s in root.findall('a:tblStyle',o.NS) if s.attrib['styleId']==c['styleId']).find('a:wholeTbl/a:tcTxStyle/a:font',o.NS)
   font=next(f for f in node.findall('a:font',o.NS) if f.attrib['script']==c['script']);font.set('typeface','Wrong')
  file=temp/(id+'-rule.pptx');rewrite(file,row['selector']['sourcePart'],rule);reject(id+'/source-rule',lambda:o.inspect_package(file,e))
  file=temp/(id+'-body.pptx');rewrite(file,part,lambda r:table(r).find('.//a:bodyPr',o.NS).set('lIns','999'),'table-fonts-round1');reject(id+'/unexpected-native-normalization',lambda:normalized(file))
  doc=fitz.open(p/'table-fonts.pdf');traces=[t for t in doc[index].get_texttrace() if t['bbox'][1]>60 and any(c[0]!=32 for c in t['chars'])];assert traces
  for t in traces:
   b=t['bbox'];doc[index].draw_rect(fitz.Rect(b[0]-2,b[1]-2,b[2]+2,b[3]+2),color=None,fill=(1,1,1),overlay=True)
  file=temp/(id+'-invisible.pdf');doc.save(file);doc.close();reject(id+'/extractable-invisible',lambda:o.inspect_pdf(file,e))
  doc=fitz.open(p/'table-fonts.pdf');actual=traces[0]['font']
  for font in doc[index].get_fonts():
   if font[3].split('+')[-1]==actual:doc.xref_set_key(font[0],'BaseFont','/Wrong')
  file=temp/(id+'-resource.pdf');doc.save(file);doc.close();reject(id+'/actual-resource',lambda:o.inspect_pdf(file,e))
 doc=fitz.open(p/'table-fonts.pdf');doc.select([1,0]+list(range(2,len(doc))));file=temp/'reordered.pdf';doc.save(file);doc.close();reject('actual-PDF-order',lambda:o.inspect_pdf(file,e))
print(json.dumps({'rejectedCount':len(rejected),'rejectedMutations':rejected}))
