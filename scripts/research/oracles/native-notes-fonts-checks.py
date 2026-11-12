"""Real notes ZIP/PDF corruption controls; never overwrite native artifacts."""
import importlib.util,json,sys,tempfile,zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz
spec=importlib.util.spec_from_file_location('notes',Path(__file__).with_name('native-notes-fonts.py'));o=importlib.util.module_from_spec(spec);spec.loader.exec_module(o)
p=Path(sys.argv[1]);expected=json.loads((p/'expected.json').read_text());rejected=[]
def reject(name,fn):
    try:fn()
    except AssertionError:rejected.append(name);return
    raise AssertionError('Accepted corruption: '+name)
def rewrite(destination,source,part,change):
    with zipfile.ZipFile(source)as z,zipfile.ZipFile(destination,'w')as out:
        assert part in z.namelist()
        for entry in z.infolist():
            data=z.read(entry.filename)
            if entry.filename==part:
                root=ET.fromstring(data);change(root);data=ET.tostring(root)
            out.writestr(entry,data)
def rp(n):return o.body(n).find('p:txBody/a:p/a:r/a:rPr',o.NS)
def local(n):
    node=rp(n).find('a:ea',o.NS)
    if node is None:node=ET.SubElement(rp(n),o.A+'ea')
    node.set('typeface','Wrong')
with tempfile.TemporaryDirectory(prefix='pptkit-notes-negative-')as temp:
    temp=Path(temp);source=p/'native-round1.pptx';pdf=p/'native-round1-notes.pdf'
    positive=o.inspect_package(source,p,expected)
    for row in positive['controls']:
        index=row['slide']-1;part=row['notesPart'];label=row['id']
        changes=[('body-text',lambda n:o.body(n).find('p:txBody/a:p/a:r/a:t',o.NS).__setattr__('text','Wrong')),
                 ('language',lambda n:rp(n).set('lang','xx-XX')),('local-font',local),
                 ('wrong-local-type',lambda n:o.body(n).find('p:nvSpPr/p:nvPr/p:ph',o.NS).set('type','title')),
                 ('page-label',lambda n:o.named(n,'Notes Control Label').find('p:txBody/a:p/a:r/a:t',o.NS).__setattr__('text','Wrong'))]
        for name,change in changes:
            file=temp/'changed.pptx';rewrite(file,source,part,change);reject(label+'/'+name,lambda:o.inspect_package(file,p,expected))
        with fitz.open(pdf)as d:
            trace=next(t for t in d[index].get_texttrace()if t['bbox'][1]>410);b=trace['bbox'];d[index].draw_rect(fitz.Rect(b[0]-2,b[1]-2,b[2]+2,b[3]+2),color=None,fill=(1,1,1),overlay=True);file=temp/'hidden.pdf';d.save(file)
        reject(label+'/extractable-invisible-notes',lambda:o.inspect_pdf(file,p,expected))
        with fitz.open(pdf)as d:
            fonts=[f for f in d[index].get_fonts()if f[3].split('+')[-1]==trace['font']];assert fonts
            for f in fonts:d.xref_set_key(f[0],'BaseFont','/Wrong')
            file=temp/'font.pdf';d.save(file)
        reject(label+'/font-resource',lambda:o.inspect_pdf(file,p,expected))
    global_changes=[
        ('actual-notes-theme','ppt/theme/theme2.xml',lambda n:n.find('a:themeElements/a:fontScheme/a:minorFont/a:font',o.NS).set('typeface','Wrong')),
        ('borrow-slide-theme','ppt/notesMasters/_rels/notesMaster1.xml.rels',lambda n:n[0].set('Target','../theme/theme1.xml')),
        ('wrong-master-type','ppt/notesMasters/notesMaster1.xml',lambda n:o.body(n).find('p:nvSpPr/p:nvPr/p:ph',o.NS).set('type','title')),
        ('body-master-size','ppt/notesMasters/notesMaster1.xml',lambda n:o.body(n).find('p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr',o.NS).set('sz','1800')),
        ('wrong-actual-master','ppt/notesSlides/_rels/notesSlide1.xml.rels',lambda n:next(c for c in n if c.attrib['Type'].endswith('/notesMaster')).set('Target','../notesMasters/missing.xml')),
        ('slide-neighbor','ppt/slides/slide1.xml',lambda n:o.named(n,'SlideProbe').find('p:txBody/a:p/a:r/a:t',o.NS).__setattr__('text','Wrong'))]
    for name,part,change in global_changes:
        file=temp/'changed.pptx';rewrite(file,source,part,change)
        # A malformed/missing relationship is a rejection too; the oracle must not silently borrow a listed master.
        try:o.inspect_package(file,p,expected)
        except (AssertionError,KeyError):rejected.append(name)
        else:raise AssertionError('Accepted corruption: '+name)
    with fitz.open(pdf)as d:
        d.select([1,0]+list(range(2,len(d))));file=temp/'order.pdf';d.save(file)
    reject('page-order',lambda:o.inspect_pdf(file,p,expected))
    with fitz.open(p/'native-round3-notes.pdf')as d:
        d[0].draw_rect(fitz.Rect(450,600,480,640),color=None,fill=(0,0,0));file=temp/'raster.pdf';d.save(file)
    reject('later-round-raster',lambda:o.inspect_pdf(file,p,expected,p/'source-notes.pdf'))
print(json.dumps({'rejectedCount':len(rejected),'rejectedMutations':rejected}))
