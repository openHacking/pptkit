"""Mutate real theme artifacts; strict baseline/source/Office/PDF checks must reject."""
import importlib.util, json, re, sys, tempfile, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz
spec = importlib.util.spec_from_file_location('table_theme', Path(__file__).with_name('native-table-font-themes.py'))
o = importlib.util.module_from_spec(spec); spec.loader.exec_module(o)
p=Path(sys.argv[1]);expected=json.loads((p/'expected.json').read_text());rejected=[]

def reject(name,fn):
    try:fn()
    except AssertionError:rejected.append(name);return
    raise AssertionError('Accepted corruption: '+name)

def rewrite(destination,source,part,change):
    with zipfile.ZipFile(source)as before,zipfile.ZipFile(destination,'w')as after:
        assert part in before.namelist()
        for entry in before.infolist():
            data=before.read(entry.filename)
            if entry.filename==part:
                root=ET.fromstring(data);change(root);data=ET.tostring(root)
            after.writestr(entry,data)

def props(root):return root.find('.//a:tbl/a:tr/a:tc/a:txBody/a:p/a:r/a:rPr',o.NS)

def replace_local(root):
    run=props(root);node=run.find('a:ea',o.NS)
    if node is None:node=ET.SubElement(run,'{'+o.NS['a']+'}ea')
    node.set('typeface','Wrong')

with tempfile.TemporaryDirectory(prefix='pptkit-table-theme-negative-')as temp:
    temp=Path(temp);source=p/'office-round1.pptx';pdf=p/'office-round1.pdf';verified=o.inspect_package(source,p,expected)
    for row in verified['controls']:
        index=row['slide']-1;label=row['id'];part=row['part']
        changes=[('text',lambda r:r.find('.//a:tbl/a:tr/a:tc/a:txBody/a:p/a:r/a:t',o.NS).__setattr__('text','Wrong')),
                 ('language',lambda r:props(r).set('lang','xx-XX')),('local-font',replace_local),
                 ('style-binding',lambda r:r.find('.//a:tbl/a:tblPr/a:tableStyleId',o.NS).__setattr__('text','Wrong'))]
        for name,change in changes:
            file=temp/'changed.pptx';rewrite(file,source,part,change);reject(label+'/'+name,lambda:o.inspect_package(file,p,expected))
        with fitz.open(pdf)as doc:
            trace=o.append.body_traces(doc[index])[-1];b=trace['bbox'];doc[index].draw_rect(fitz.Rect(b[0]-2,b[1]-2,b[2]+2,b[3]+2),color=None,fill=(1,1,1),overlay=True);file=temp/'hidden.pdf';doc.save(file)
        reject(label+'/extractable-invisible-body-or-omega',lambda:o.inspect_pdf(file,p,expected))
        with fitz.open(pdf)as doc:
            fonts=[f for f in doc[index].get_fonts()if f[3].split('+')[-1]==trace['font']];assert fonts
            for font in fonts:doc.xref_set_key(font[0],'BaseFont','/Wrong')
            file=temp/'font.pdf';doc.save(file)
        reject(label+'/font-resource',lambda:o.inspect_pdf(file,p,expected))
    package_changes=[
        ('duplicate-content-type','[Content_Types].xml',lambda r:r.append(ET.fromstring(ET.tostring(next(n for n in r if n.tag.endswith('}Override')))))),
        ('theme-slot','ppt/theme/theme1.xml',lambda r:r.find('a:themeElements/a:fontScheme/a:minorFont/a:latin',o.NS).set('typeface','Wrong')),
        ('layout-binding','ppt/slides/_rels/slide1.xml.rels',lambda r:next(n for n in r if n.attrib['Type'].endswith('/slideLayout')).set('Target','../slideLayouts/slideLayout1.xml')),
        ('theme-binding','ppt/slideMasters/_rels/slideMaster1.xml.rels',lambda r:next(n for n in r if n.attrib['Type'].endswith('/theme')).set('Target','../theme/theme2.xml')),
        ('office-layout-semantic','ppt/slideLayouts/slideLayout7.xml',lambda r:r.set('type','title')),
        ('office-creation-guid-format','ppt/slideMasters/slideMaster1.xml',lambda r:next(r.iter('{'+o.A16+'}creationId')).set('id','bad')),
        ('office-creation-guid-duplicate','ppt/slideMasters/slideMaster1.xml',lambda r:list(r.iter('{'+o.A16+'}creationId'))[1].set('id',list(r.iter('{'+o.A16+'}creationId'))[0].attrib['id'])),
        ('office-shared-field-identity','ppt/slideMasters/slideMaster1.xml',lambda r:next(r.iter('{'+o.NS['a']+'}fld')).set('id','{11111111-1111-1111-1111-111111111111}')),
        ('office-creation-number','ppt/slideMasters/slideMaster1.xml',lambda r:next(r.iter('{'+o.append.P14+'}creationId')).set('val',str(2**32))),
        ('unrelated-notes','ppt/notesSlides/notesSlide1.xml',lambda r:r.set('showMasterSp','0'))]
    for name,part,change in package_changes:
        file=temp/'changed.pptx';rewrite(file,source,part,change);reject(name,lambda:o.inspect_package(file,p,expected))
    with fitz.open(pdf)as doc:
        doc.select([1,0]+list(range(2,len(doc))));file=temp/'order.pdf';doc.save(file)
    reject('page-order',lambda:o.inspect_pdf(file,p,expected))
    with fitz.open(pdf)as doc:
        changed=0
        for xref in doc[0].get_contents():
            data=doc.xref_stream(xref).decode('latin1')
            def shift(m):
                global changed
                changed+=1;return m.group(1)+str(float(m.group(2))+10)+' Tm'
            data=re.sub(r'(BT [^\r\n]*? )(-?[0-9.]+) Tm',shift,data);doc.update_stream(xref,data.encode('latin1'))
        assert changed;file=temp/'position.pdf';doc.save(file)
    reject('unexpected-label-baseline',lambda:o.inspect_pdf(file,p,expected))
    with fitz.open(p/'office-round3.pdf')as doc:
        doc[0].draw_rect(fitz.Rect(500,180,530,200),color=None,fill=(0,0,0));file=temp/'round.pdf';doc.save(file)
    reject('later-round-raster',lambda:o.inspect_pdf(file,p,expected,previous=pdf))
print(json.dumps({'rejectedCount':len(rejected),'rejectedMutations':rejected}))
