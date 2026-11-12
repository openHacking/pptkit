"""Independent native master artwork ownership/visibility/geometry checks for six named pages."""
import collections
import copy
import hashlib
import importlib.util
import io
import json
import re
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
import xml.etree.ElementTree as E
import fitz
from PIL import Image
HERE=Path(__file__).resolve().parent

def load(name,file):
    spec=importlib.util.spec_from_file_location(name,file); module=importlib.util.module_from_spec(spec); spec.loader.exec_module(module); return module
n=load('master_native_primitives',HERE/'multi-master-native.py')
m=load('master_font_primitives',HERE/'multi-master-operations.py')
P,A,R,NS=n.P,n.A,n.R,n.NS
sha=n.sha
META=['ppt/presProps.xml','ppt/viewProps.xml','docProps/core.xml','docProps/app.xml']
EXPECTED=[('red','red-on',True,True),('red','red-on',False,False),('red','red-off',False,True),('blue','blue-on',True,True),('red','red-off',False,True),('red','red-on',True,True)]

def graph(parts):
    main=n.target(parts,'','officeDocument'); root=n.root(parts,main); aliases=collections.defaultdict(set)
    def add(label,part): assert part in parts; aliases[part].add(label)
    add('presentation',main); add('theme/presentation',n.target(parts,main,'theme')); edges=n.relations(parts,main); masters={}; layouts={}
    for ref in root.findall('p:sldMasterIdLst/p:sldMasterId',NS):
        id=ref.get('id'); part=edges[ref.get('{'+R+'}id')]['target']; assert id not in masters; masters[id]=part; add('master/'+id,part);add('theme/master/'+id,n.target(parts,part,'theme'))
        for lr in n.root(parts,part).findall('p:sldLayoutIdLst/p:sldLayoutId',NS):
            lid=lr.get('id'); lp=n.relations(parts,part)[lr.get('{'+R+'}id')]['target'];assert lid not in layouts and n.target(parts,lp,'slideMaster')==part;layouts[lid]=lp;add('layout/'+lid,lp)
    assert len(masters)==3 and len(layouts)==4 and set(masters).isdisjoint(layouts),'Owner graph or numeric namespace changed'
    refs=root.findall('p:notesMasterIdLst/p:notesMasterId',NS);assert len(refs)==1;nm=edges[refs[0].get('{'+R+'}id')]['target'];add('notes-master',nm);add('theme/notes',n.target(parts,nm,'theme'))
    rows=[]
    for index,ref in enumerate(root.findall('p:sldIdLst/p:sldId',NS)):
        id=ref.get('id');part=edges[ref.get('{'+R+'}id')]['target'];add('slide/'+id,part)
        lp=n.target(parts,part,'slideLayout');mp=n.target(parts,lp,'slideMaster');note=n.target(parts,part,'notesSlide');add('notes/'+id,note)
        assert n.target(parts,note,'slide')==part and n.target(parts,note,'notesMaster')==nm
        rows.append({'page':index,'slide':part,'layout':lp,'master':mp,'theme':n.target(parts,mp,'theme'),'notes':note})
    assert len(rows)==6
    media=[part for part in parts if part.startswith('ppt/media/')];assert len(media)==1,'Shared image was duplicated/lost';add('media/'+sha(parts[media[0]]),media[0])
    for part in list(aliases):
        if not part.startswith('ppt/media/'):
            ids=[node.get('id') for node in n.root(parts,part).iter('{'+P+'}cNvPr')];assert len(set(ids))==len(ids) and all(id.isdigit() and int(id)>0 for id in ids),'Part-local drawing IDs collided'
        for edge in n.relations(parts,part).values():
            if edge['mode']=='Internal':assert edge['target'] in parts,'Broken relationship closure'
    return aliases,rows

def semantic(parts):
    aliases,rows=graph(parts);label=lambda part:'|'.join(sorted(aliases.get(part,{part})))
    def canonical(node,part,parent=None):
        attrs=dict(node.attrib)
        # Native adds this named preset to the exact 720x540 source frame.
        if node.tag=='{'+P+'}sldSz' and attrs=={'cx':'9144000','cy':'6858000','type':'screen4x3'}:del attrs['type']
        if node.tag=='{'+A+'}pPr' and attrs.get('indent')=='-0':attrs['indent']='0'
        if node.tag in ['{'+P+'}sld','{'+P+'}sldLayout'] and attrs.get('showMasterSp')=='1':del attrs['showMasterSp']
        for key,value in list(attrs.items()):
            if key.startswith('{'+R+'}'):
                edge=n.relations(parts,part)[value];attrs[key]=(edge['type'],edge['mode'],label(edge['target']) if edge['mode']=='Internal' else edge['target'])
        children=[]
        for child in node:
            if node.tag=='{'+P+'}cNvSpPr' and child.tag=='{'+A+'}spLocks' and len(child)==0:
                ph=parent.find('p:nvPr/p:ph',NS) if parent is not None else None
                if child.attrib=={'noGrp':'1'} or (ph is not None and ph.get('type')=='sldImg' and child.attrib=={'noGrp':'1','noRot':'1','noChangeAspect':'1'}):continue
            if node.tag=='{'+P+'}presentation' and child.tag=='{'+P+'}extLst' and n.guide_extension(child):continue
            children.append(canonical(child,part,node))
        return node.tag,sorted(attrs.items()),node.text or '',children
    types=n.root(parts,'[Content_Types].xml');result={'content-types':sorted((c.tag,c.get('Extension') if c.get('Extension') is not None else label(c.get('PartName').lstrip('/')),c.get('ContentType')) for c in types)}
    def rels(part):return sorted((e['type'],e['mode'],label(e['target']) if e['mode']=='Internal' else e['target']) for e in n.relations(parts,part).values())
    result['root-relationships']=rels('');covered={'[Content_Types].xml','_rels/.rels'}
    for part in aliases:
        result[label(part)]=sha(parts[part]) if part.startswith('ppt/media/') else canonical(n.root(parts,part),part);result[label(part)+'/relationships']=rels(part);covered.update([part,n.relname(part)])
    for part,body in parts.items():
        if part in covered or part in META:continue
        result[part]=n.plain(n.root(parts,part)) if part.endswith('.xml') else sha(body)
    return result,rows

def same(actual,expected):
    assert actual.keys()==expected.keys(),'Lost/extra semantic entries';different=[key for key in actual if actual[key]!=expected[key]];assert not different,('Unapproved package changes',different)

def package(file,expected):
    parts=n.payload(file);actual,rows=semantic(parts);same(actual,expected)
    for index,row in enumerate(rows):
        master,layout,show_m,show_l=EXPECTED[index];mr=n.root(parts,row['master']);lr=n.root(parts,row['layout']);sr=n.root(parts,row['slide'])
        assert mr.find('p:cSld',NS).get('name')==master and lr.find('p:cSld',NS).get('name')==layout,'Crossed owner/page binding'
        assert (sr.get('showMasterSp','1')!='0')==show_l
        assert ((lr.get('showMasterSp','1')!='0') and show_l)==show_m
        for owner,prefix in [(row['master'],master+'-master-'),(row['layout'],layout+'-layout-')]:
            refs=[e for e in n.relations(parts,owner).values() if e['type'].endswith('/image')];assert len(refs)==1
            assert refs[0]['mode']=='Internal' and parts[refs[0]['target']]==parts[next(p for p in parts if p.startswith('ppt/media/'))]
            assert not any(e['type'].endswith('/image') for e in n.relations(parts,row['slide']).values()),'Artwork copied to slide'
            assert any(c.get('name','').startswith(prefix) for c in n.root(parts,owner).iter('{'+P+'}cNvPr'))
        assert len(lr.findall('.//p:ph',NS))==1 and len(mr.findall('.//p:ph',NS))==0,'Artwork acquired prompt identity'
        assert sr.find('p:cSld/p:spTree',NS) is not None
    return parts,rows

def metadata(parts):
    result={}
    for part in META:
        root=n.root(parts,part)
        if part=='docProps/core.xml':
            assert root.find('{http://purl.org/dc/elements/1.1/}title').text=='Master artwork controls'
            assert root.find('{http://purl.org/dc/elements/1.1/}creator').text=='PPTKit'
            for child in root:
                if child.tag.split('}')[-1] in ['modified','lastPrinted']:
                    datetime.fromisoformat(child.text.replace('Z','+00:00'));child.text='native-clock'
        if part=='docProps/app.xml':
            for child in root:
                if child.tag.endswith('}TotalTime'):assert child.text.isdigit();child.text='native-total-time'
        result[part]=n.plain(root)
    return result

def plain(value):
    if isinstance(value,dict):return {k:plain(v) for k,v in value.items()}
    if isinstance(value,(list,tuple,fitz.Rect,fitz.Point,fitz.Quad)):return [plain(v) for v in value]
    if isinstance(value,bytes):return value.hex()
    return value

def rgb(value):return tuple(round(v*255) for v in value)
def pdf(file,baseline=None):
    doc=fitz.open(file);assert len(doc)==6,'Page count changed';other=fitz.open(baseline) if baseline else None;result=[]
    for i,page in enumerate(doc):
        master,layout,show_m,show_l=EXPECTED[i]
        texts=([f'MASTER {master}'] if show_m else [])+([f'LAYOUT {layout}'] if show_l else [])+[f'PAGE {i}',f'BODY {i}']
        assert page.get_text().strip().splitlines()==texts,'Hidden/lost/reordered artwork or page correspondence'
        spans=[s for b in page.get_text('dict')['blocks'] for l in b.get('lines',[]) for s in l['spans']];assert all(s['font']=='ArialMT' for s in spans),'Unexpected actual face'
        drawings=page.get_drawings();fills=[d for d in drawings if d['fill'] is not None]
        colors=collections.Counter(rgb(d['fill']) for d in fills)
        accent=(204,34,0) if master=='red' else (0,68,204);background=(255,255,221) if i==5 else (255,248,240) if master=='red' else (240,248,255)
        accent_int=(accent[0]<<16)|(accent[1]<<8)|accent[2];assert [(s['text'],s['color']) for s in spans]==[(text,0 if text.startswith('BODY ') else accent_int) for text in texts],'Wrong actual text color/domain'
        assert colors[accent]==int(show_m) and colors[(136,68,170)]==int(show_m) and colors[(0,170,85)]==int(show_l) and colors[(255,204,0)]==1,'Artwork fills/visibility changed'
        frame=next(d['rect'] for d in fills if rgb(d['fill'])==background)
        sx,sy=frame.width/720,frame.height/540
        def geometry(rect,source):
            expected=[frame.x0+source[0]*sx,frame.y0+source[1]*sy,frame.x0+source[2]*sx,frame.y0+source[3]*sy]
            assert all(abs(a-b)<0.0001 for a,b in zip(rect,expected)),('Source-frame geometry mismatch',rect,expected)
        geometry(next(d['rect'] for d in fills if rgb(d['fill'])==(255,204,0)),[100,60,160,140])
        if show_m:
            geometry(next(d['rect'] for d in fills if rgb(d['fill'])==accent),[20,20,280,100]);geometry(next(d['rect'] for d in fills if rgb(d['fill'])==(136,68,170)),[500,20,550,70])
        if show_l:geometry(next(d['rect'] for d in fills if rgb(d['fill'])==(0,170,85)),[60,40,200,120])
        images=page.get_image_info(hashes=True);assert len(images)==int(show_m)+int(show_l),'Shared image appearance count changed'
        sources=([[20,240,84,272]] if show_m else [])+([[300,240,364,272]] if show_l else [])
        for image,source in zip(images,sources):assert image['width']==8 and image['height']==4 and image['digest'].hex()=='b851ddcf2b138091f4b239006e0fea8f';geometry(image['bbox'],source)
        pix=page.get_pixmap(matrix=fitz.Matrix(2,2),alpha=False)
        for x,y,color in [(40,40,accent if show_m else background),(80,60,(0,170,85) if show_l else accent if show_m else background),(120,80,(255,204,0))]:
            actual=pix.pixel(round((frame.x0+x*sx)*2),round((frame.y0+y*sy)*2));assert all(abs(a-b)<=1 for a,b in zip(actual,color)),'Wrong physical layer order/color'
        record={'page':i,'size':list(page.rect),'text':page.get_text(),'rasterSha256':sha(pix.samples),'trace':plain(page.get_texttrace()),'drawings':plain(drawings),'images':plain(images),'fonts':m.fonts(doc,page)}
        if other:
            original=pdf_page_record(other,i);same(record,original)
        result.append(record)
    return result

def pdf_page_record(doc,i):
    page=doc[i];return {'page':i,'size':list(page.rect),'text':page.get_text(),'rasterSha256':sha(page.get_pixmap(matrix=fitz.Matrix(2,2),alpha=False).samples),'trace':plain(page.get_texttrace()),'drawings':plain(page.get_drawings()),'images':plain(page.get_image_info(hashes=True)),'fonts':m.fonts(doc,page)}

def negatives(directory,expected):
    out=Path(tempfile.mkdtemp(prefix='negative-',dir=directory));base=n.payload(directory/'artwork-round1.pptx');_,rows=graph(base);results=[]
    def xml_patch(part,fn):
        def mutate(parts):
            root=n.root(parts,part);fn(root);parts[part]=E.tostring(root,encoding='utf-8',xml_declaration=True)
        return mutate
    def object_named(root,name):
        return next(node for node in root.iter('{'+P+'}sp') if node.find('p:nvSpPr/p:cNvPr',NS).get('name')==name)
    def drop_image(root):root.find('p:cSld/p:spTree',NS).remove(root.find('p:cSld/p:spTree/p:pic',NS))
    def duplicate_to_slide(parts):
        root=n.root(parts,rows[0]['slide']);node=copy.deepcopy(object_named(n.root(parts,rows[0]['master']),'red-master-shape'));node.find('p:nvSpPr/p:cNvPr',NS).set('id','99');root.find('p:cSld/p:spTree',NS).append(node);parts[rows[0]['slide']]=E.tostring(root,encoding='utf-8',xml_declaration=True)
    def wrong_picture(parts):
        part=next(p for p in parts if p.startswith('ppt/media/'));image=Image.open(io.BytesIO(parts[part])).convert('RGB');image.putpixel((0,0),(0,0,0));buffer=io.BytesIO();image.save(buffer,format='PNG');parts[part]=buffer.getvalue()
    def break_image(root):next(edge for edge in root if edge.get('Type').endswith('/image')).set('Target','../media/missing.png')
    def wrong_owner(root):next(edge for edge in root if edge.get('Type').endswith('/slideMaster')).set('Target','../slideMasters/slideMaster3.xml')
    def wrong_group(root):next(node for node in root.iter('{'+A+'}chExt') if node.get('cx')=='1270000').set('cx','1270001')
    def cross_theme(root):next(edge for edge in root if edge.get('Type').endswith('/theme')).set('Target','../theme/theme4.xml')
    def collision(root):root.find('p:sldLayoutIdLst/p:sldLayoutId',NS).set('id','2147483648')
    def prompt(root):E.SubElement(object_named(root,'red-master-shape').find('p:nvSpPr/p:nvPr',NS),'{'+P+'}ph',{'type':'body','idx':'1'})
    def text_color(root):object_named(root,'red-master-text').find('p:txBody/a:p/a:r/a:rPr/a:solidFill/a:schemeClr',NS).set('val','accent2')
    def page_order(root):lst=root.find('p:sldIdLst',NS);a,b=list(lst)[:2];lst.remove(b);lst.insert(0,b)
    cases=[('wrong-master-text-color',xml_patch(rows[0]['master'],text_color)),('layout-owner-crossing',xml_patch(n.relname(rows[0]['layout']),wrong_owner)),('copied-master-shape-on-slide',duplicate_to_slide),('lost-master-picture',xml_patch(rows[0]['master'],drop_image)),('changed-picture-pixels',wrong_picture),('broken-image-closure',xml_patch(n.relname(rows[0]['master']),break_image)),('layout-show-lost',xml_patch(rows[2]['layout'],lambda root:root.attrib.pop('showMasterSp'))),('slide-show-lost',xml_patch(rows[1]['slide'],lambda root:root.attrib.pop('showMasterSp'))),('group-one-emu',xml_patch(rows[0]['master'],wrong_group)),('master-theme-crossing',xml_patch(n.relname(rows[0]['master']),cross_theme)),('id-namespace-collision',xml_patch(rows[0]['master'],collision)),('master-became-prompt',xml_patch(rows[0]['master'],prompt)),('page-order-swapped',xml_patch('ppt/presentation.xml',page_order)),('wrong-size-preset',xml_patch('ppt/presentation.xml',lambda root:root.find('p:sldSz',NS).set('type','screen16x9'))),('changed-creator-title',xml_patch('docProps/core.xml',lambda root:root.find('{http://purl.org/dc/elements/1.1/}title').__setattr__('text','Wrong')))]
    for name,mutate in cases:
        parts=copy.copy(base);mutate(parts);file=out/(name+'.pptx')
        with zipfile.ZipFile(file,'w',compression=zipfile.ZIP_DEFLATED) as z:
            for part,body in parts.items():z.writestr(part,body)
        try:parts,_=package(file,expected);metadata(parts)
        except (AssertionError,KeyError,StopIteration,TypeError,ValueError) as error:results.append({'name':name,'kind':'zip','file':str(file),'sha256':sha(file.read_bytes()),'reason':str(error)})
        else:raise AssertionError('Accepted actual corrupt package '+name)
    def reject_pdf(name,file):
        try:pdf(file,directory/'first.pdf')
        except (AssertionError,KeyError,StopIteration) as error:results.append({'name':name,'kind':'pdf','file':str(file),'sha256':sha(file.read_bytes()),'reason':str(error)})
        else:raise AssertionError('Accepted actual bad PDF '+name)
    for name in ['covered-master','swapped-pages','covered-picture']:
        doc=fitz.open(directory/'artwork-round1.pdf');file=out/(name+'.pdf')
        if name=='covered-master':doc[0].draw_rect(doc[0].search_for('MASTER red')[0],fill=(1,1,1),color=(1,1,1),overlay=True)
        elif name=='covered-picture':doc[0].draw_rect(fitz.Rect(doc[0].get_image_info()[0]['bbox']),fill=(1,1,1),color=(1,1,1),overlay=True)
        else:doc.select([1,0,2,3,4,5])
        doc.save(file);reject_pdf(name,file)
    doc=fitz.open(directory/'artwork-round1.pdf');page=doc[0];font=next(f for f in page.get_fonts() if f[3].split('+')[-1]=='ArialMT');program=doc.extract_font(font[0])[3];gid=page.get_texttrace()[0]['chars'][0][1]
    altered=m.alter_font_advance(program,gid);before=m.glyph_program(program,gid);after=m.glyph_program(altered,gid);assert after==(before[0],before[1]+1,before[2],before[3])
    streams=[xref for xref in range(1,doc.xref_length()) if doc.xref_is_stream(xref) and doc.xref_stream(xref)==program];assert len(streams)==1
    doc.update_stream(streams[0],altered);file=out/'same-raster-font-advance.pdf';doc.save(file);patched=fitz.open(file);positive=fitz.open(directory/'artwork-round1.pdf');assert all(patched[i].get_pixmap(matrix=fitz.Matrix(2,2),alpha=False).samples==positive[i].get_pixmap(matrix=fitz.Matrix(2,2),alpha=False).samples for i in range(6))
    reject_pdf('same-raster-font-advance',file);results[-1].update({'sameRaster':True,'metricBefore':before,'metricAfter':after})
    return results

def main():
    manifest=json.loads(Path('docs/testing/master-artwork-evidence.json').read_text());directory=Path(manifest['nativeDirectory']);original=n.payload(directory/'input.pptx');expected,_=semantic(original)
    assert sha((directory/'input.pptx').read_bytes())==manifest['sourceSha256'];package(directory/'input.pptx',expected)
    opened=(directory/'first-open.ax.txt').read_text();assert 'URL: file://'+str(directory/'input.pptx') in opened and not re.search('发现.*内容有问题|修复',opened)
    media=next(part for part in original if part.startswith('ppt/media/'));source_pixels=Image.open(io.BytesIO(original[media])).convert('RGB').tobytes();assert hashlib.md5(source_pixels).hexdigest()=='b851ddcf2b138091f4b239006e0fea8f'
    first=pdf(directory/'first.pdf');rounds=[];first_meta=None
    for i in [1,2,3]:
        file=directory/f'artwork-round{i}.pptx';parts,rows=package(file,expected);meta=metadata(parts)
        if first_meta is None:first_meta=meta
        else:same(meta,first_meta)
        observations=pdf(directory/f'artwork-round{i}.pdf',directory/'first.pdf')
        for phase in ['saved','reopened']:
            ax=(directory/f'artwork-round{i}-{phase}.ax.txt').read_text();assert 'URL: file://'+str(file) in ax and not re.search('发现.*内容有问题|修复',ax),'Native identity or repair prompt'
        assert str(file) not in (directory/f'artwork-round{i}-closed.ax.txt').read_text(),'Document was not closed'
        rounds.append({'round':i,'file':str(file),'sha256':sha(file.read_bytes()),'metadata':{p:parts[p].decode() for p in META},'pages':observations})
    about=(directory/'powerpoint-version.ax.txt').read_text();assert '16.104 (25121423)' in about
    bad=negatives(directory,expected)
    # A compact contact sheet is a visual supplement, not the comparison oracle.
    contact=Image.new('RGB',(3*421,2*298),(224,224,224));doc=fitz.open(directory/'first.pdf')
    for i in range(6):
        pix=doc[i].get_pixmap(matrix=fitz.Matrix(.5,.5),alpha=False);image=Image.frombytes('RGB',[pix.width,pix.height],pix.samples);contact.paste(image,((i%3)*421,(i//3)*298))
    contact.save(directory/'native-contact.png')
    files=[]
    for file in sorted(directory.rglob('*')):
        if file.is_file():body=file.read_bytes();files.append({'file':str(file),'bytes':len(body),'sha256':sha(body)})
    deps=[{'file':str(file),'sha256':sha(file.read_bytes())} for file in [Path(__file__),HERE/'multi-master-native.py',HERE/'multi-master-operations.py']]
    result={'schemaVersion':1,'recordedAt':datetime.now(timezone.utc).isoformat(),'directory':str(directory),'sourceSha256':manifest['sourceSha256'],'application':'PowerPoint 16.104 (25121423)','verified':True,'firstPages':first,'rounds':rounds,'negatives':bad,'dependencies':deps,'artifacts':files,'firstSaveMetadata':{p:{'source':original[p].decode(),'native':n.payload(directory/'artwork-round1.pptx')[p].decode()} for p in META},'namedNormalizations':['media filename aliases by exact payload SHA; one shared image only','showMasterSp=1 schema default omission','screen4x3 preset addition only on exact 720x540 frame','pPr negative zero','named noGrp/notes image locks and empty slide guides'],'scope':'Six named shape/text/image/group artwork pages, source-frame affine geometry, actual layer colors/pixels, complete resource/part graph, native 3 save-close-exact-reopen rounds. Native table/chart-master appearance, shared editing and broader source import/style writes remain separate.'}
    Path('docs/testing/master-artwork-native-evidence.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps({'verified':True,'firstPages':len(first),'reopenedPages':sum(len(r['pages']) for r in rounds),'rounds':len(rounds),'zipNegatives':sum(n['kind']=='zip' for n in bad),'pdfNegatives':sum(n['kind']=='pdf' for n in bad),'artifacts':len(files)}))

if __name__=='__main__':main()
