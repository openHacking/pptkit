"""Independent notes domain, named native normalization and actual notes-page PDF oracle."""
import copy, datetime, hashlib, importlib.util, json, math, sys, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz
from PIL import Image
spec=importlib.util.spec_from_file_location('table',Path(__file__).with_name('native-table-fonts.py'))
table=importlib.util.module_from_spec(spec);spec.loader.exec_module(table)
NS=table.NS;A='{'+NS['a']+'}';P='{'+NS['p']+'}'
REL='http://schemas.openxmlformats.org/package/2006/relationships'

def sha(file):return hashlib.sha256(Path(file).read_bytes()).hexdigest()
def tree(n):return(n.tag,sorted(n.attrib.items()),n.text or None,[tree(c)for c in n])
def digest(n):return hashlib.sha256(json.dumps(tree(n),ensure_ascii=False).encode()).hexdigest()
def relation(z,owner,kind,id=None):return table.mixed.edit.relationship(z,owner,kind,id)
def root(z,part):return ET.fromstring(z.read(part))
def named(n,name):
    found=[s for s in n.findall('p:cSld/p:spTree/p:sp',NS)if s.find('p:nvSpPr/p:cNvPr',NS).attrib['name']==name];assert len(found)==1;return found[0]
def body(n):
    found=[s for s in n.findall('p:cSld/p:spTree/p:sp',NS)if s.find('p:nvSpPr/p:nvPr/p:ph',NS)is not None and s.find('p:nvSpPr/p:nvPr/p:ph',NS).attrib['type']=='body'];assert len(found)==1;return found[0]
def normalize_paragraphs(n):
    for p in n.findall('.//a:pPr',NS):
        if p.attrib.get('indent')=='-0':p.set('indent','0')
def add_locks(n):
    for shape in n.findall('p:cSld/p:spTree/p:sp',NS):
        ph=shape.find('p:nvSpPr/p:nvPr/p:ph',NS)
        if ph is None:continue
        typ=ph.attrib['type'];attrs={'noGrp':'1'}
        if typ=='sldImg':attrs.update({'noRot':'1','noChangeAspect':'1'})
        else:assert typ=='body'
        parent=shape.find('p:nvSpPr/p:cNvSpPr',NS);assert len(parent)==0;ET.SubElement(parent,A+'spLocks',attrs)

def inspect_package(file,directory,expected,source=False):
    assert sha(directory/'notes-fonts.pptx')==expected['sourceSha256']
    assert sha(directory/'author-base.pptx')==expected['authorBaseSha256']
    if source:assert sha(file)==expected['sourceSha256']
    with zipfile.ZipFile(directory/'notes-fonts.pptx')as original,zipfile.ZipFile(file)as z:
        assert len(z.namelist())==len(set(z.namelist()))
        removed=set()if source else{expected['themePart']};assert set(z.namelist())==set(original.namelist())-removed,'part inventory'
        presentation=root(z,'ppt/presentation.xml');slides=presentation.findall('p:sldIdLst/p:sldId',NS);assert len(slides)==15
        declared_master=relation(z,'ppt/presentation.xml','notesMaster',presentation.find('p:notesMasterIdLst/p:notesMasterId',NS).attrib['{'+NS['r']+'}id'])
        notes_master='ppt/notesMasters/notesMaster1.xml';theme_part=expected['themePart']if source else'ppt/theme/theme2.xml';assert declared_master==notes_master
        assert relation(z,notes_master,'theme')==theme_part
        scheme=root(z,theme_part);assert scheme.attrib['name']=='PPTKit Notes '+expected['variant']
        for kind in ['major','minor']:assert table.collection(scheme.find('a:themeElements/a:fontScheme/a:'+kind+'Font',NS))=={'slots':[expected['map']['Latn'],'',''],'rules':expected['map']}
        assert tree(scheme)==tree(root(original,expected['themePart'])),'entire independent notes theme changed'
        master=root(z,notes_master);master_body=body(master);assert master_body.find('p:nvSpPr/p:nvPr/p:ph',NS).attrib=={'type':'body','idx':'99'}
        expected_master=root(original,notes_master)
        if not source:
            add_locks(expected_master);rpr=expected_master.find('p:notesStyle/a:lvl1pPr/a:defRPr',NS);paint=rpr.find('a:solidFill',NS);assert tree(paint)==tree(ET.fromstring('<a:solidFill xmlns:a="'+NS['a']+'"><a:srgbClr val="000000"/></a:solidFill>'));rpr.remove(paint)
        assert tree(master)==tree(expected_master),'notes master changed beyond named locks/redundant black'
        rows=[];semantic={notes_master:digest(master),theme_part:digest(scheme)}
        for i,(slide,c)in enumerate(zip(slides,expected['controls']),1):
            part=relation(z,'ppt/presentation.xml','slide',slide.attrib['{'+NS['r']+'}id']);assert part==f'ppt/slides/slide{i}.xml'
            notes=relation(z,part,'notesSlide');assert notes==f'ppt/notesSlides/notesSlide{i}.xml'and relation(z,notes,'slide')==part and relation(z,notes,'notesMaster')==notes_master
            layout=relation(z,part,'slideLayout');sm=relation(z,layout,'slideMaster');st=relation(z,sm,'theme');assert st=='ppt/theme/theme1.xml'and st!=theme_part
            sld=root(z,part);before=root(original,part)
            if not source:normalize_paragraphs(before)
            assert tree(sld)==tree(before),'slide changed beyond signed zero'
            assert named(sld,'SlideLabel').find('p:txBody/a:p/a:r/a:t',NS).text==f"notes-{expected['variant']}-{c['id']}"
            n=root(z,notes);before=root(original,notes)
            if not source:
                normalize_paragraphs(before);add_locks(before)
                if c['id']=='latin':
                    paragraph=body(before).find('p:txBody/a:p',NS);end=paragraph.find('a:endParaRPr',NS);assert end.attrib=={'lang':'en-US'};paragraph.remove(end)
            assert tree(n)==tree(before),'notes changed beyond named signed zero/locks/Latin end properties'
            b=body(n);ph=b.find('p:nvSpPr/p:nvPr/p:ph',NS);assert ph.attrib=={'type':'body','idx':'7'}
            run=b.findall('p:txBody/a:p/a:r',NS);assert len(run)==1;run=run[0];rp=run.find('a:rPr',NS);assert rp.attrib=={'lang':c['language']}and run.find('a:t',NS).text==c['text']
            fonts=[rp.find('a:'+slot,NS)for slot in ['latin','ea','cs']];fonts=[f.attrib['typeface']if f is not None else None for f in fonts];assert fonts==c['localSlots'],'local declarations'
            label=named(n,'Notes Control Label').find('p:txBody/a:p/a:r/a:t',NS).text;assert label==f"notes-{expected['variant']}-{c['id']}"
            rows.append({'id':c['id'],'slide':i,'slidePart':part,'notesPart':notes,'notesMaster':notes_master,'notesTheme':theme_part,'slideTheme':st,'typeMatch':{'type':'body','localIdx':'7','masterIdx':'99'},'text':c['text'],'language':c['language'],'localSlots':fonts,'declaredRules':expected['map'],'notesTreeSha256':digest(n),'slideTreeSha256':digest(sld)})
            semantic[notes]=digest(n);semantic[part]=digest(sld)
        for part in ['ppt/slideMasters/slideMaster1.xml','ppt/slideLayouts/slideLayout1.xml','ppt/theme/theme1.xml','ppt/tableStyles.xml']:
            assert tree(root(z,part))==tree(root(original,part)),'unrelated semantic part changed';semantic[part]=digest(root(z,part))
        # All relationship edges are compared independently of application-renumbered IDs.
        for part in [p for p in z.namelist()if p.endswith('.rels')]:
            def edges(owner,rename):
                nodes=list(root(owner,part));assert len({n.attrib['Id']for n in nodes})==len(nodes)
                return sorted(tuple(sorted((k,v.replace('../theme/independent-notes-theme.xml','../theme/theme2.xml')if rename and k=='Target'else v)for k,v in n.attrib.items()if k!='Id'))for n in nodes)
            assert edges(z,False)==edges(original,not source),'relationship graph changed: '+part
        changed=[part for part in z.namelist()if part in original.namelist()and z.read(part)!=original.read(part)]
        ancillary={part:{'sourceSha256':hashlib.sha256(original.read(part)).hexdigest(),'nativeSha256':hashlib.sha256(z.read(part)).hexdigest(),'sourceXml':original.read(part).decode(),'nativeXml':z.read(part).decode()}for part in ['ppt/presentation.xml','ppt/presProps.xml','ppt/viewProps.xml','docProps/core.xml','docProps/app.xml','[Content_Types].xml']}

        return{'sha256':sha(file),'controls':rows,'semanticHashes':semantic,'changedPartBytes':changed,'removedParts':sorted(removed),'ancillaryNormalizationRecords':ancillary,'scope':'Complete selected notes/slide/theme/master/layout/style trees and all relationship edges; metadata/view/presentation extension normalization recorded separately, not general ZIP byte preservation.'}


def observed(c,variant):
    if c['id'].startswith('local-'):return'ArialUnicodeMS'
    if c['id']=='bad-ea-tc':return'MicrosoftYaHei'
    if c['id']=='empty-ea-tc':return'MicrosoftYaHei'if variant=='A'else'SimSun'
    if variant=='A':return{'latin':'ArialMT','sc':'HiraginoSansGB-W3','tc':'STSongti-TC-Regular','ja':'HiraMinProN-W3','ko':'MalgunGothic','ko-kr':'AppleSDGothicNeo-Regular','ar':'GeezaPro','he':'ArialHebrew','greek':'Georgia','major-tc':'STSongti-TC-Regular'}[c['id']]
    return{'latin':'Georgia','sc':'STSongti-TC-Regular','tc':'HiraginoSansGB-W3','ja':'HiraginoSansGB-W3','ko':'Batang','ko-kr':'ArialUnicodeMS','ar':'ArialUnicodeMS','he':'ArialUnicodeMS','greek':'ArialMT','major-tc':'HiraginoSansGB-W3'}[c['id']]

def inspect_pdf(file,directory,expected,baseline=None):
    rows=[];fonts={f['family']:f['postscriptName']for f in expected['fonts']}
    with fitz.open(file)as doc:
        assert len(doc)==15;old=fitz.open(baseline)if baseline else None
        try:
            for i,(page,c)in enumerate(zip(doc,expected['controls'])):
                assert tuple(page.rect)==(0,0,595,842),'actual notes print page size'
                traces=page.get_texttrace();headers=[t for t in traces if 350<t['bbox'][1]<410];labels=[t for t in traces if t['bbox'][1]<350];body=[t for t in traces if t['bbox'][1]>410];assert len(headers)==1 and len(body)==1
                label=f"notes-{expected['variant']}-{c['id']}";pieces=label.split('-');expected_labels=[]
                for j,piece in enumerate(pieces):
                    if j:expected_labels.append('-')
                    expected_labels.append(piece)
                assert [''.join(chr(x[0])for x in t['chars'])for t in labels]==expected_labels,'named label segmentation/order'
                assert ''.join(chr(x[0])for x in headers[0]['chars'])==f"Independent Notes {expected['variant']}",'notes domain'
                assert len(traces)==len(headers)+len(labels)+len(body),'unexpected note traces'
                assert all(t['font']=='ArialMT'and abs(t['size']-17)<.001 for t in headers+labels)
                t=body[0];face=observed(c,expected['variant']);assert t['font']==face and t['type']==0 and abs(t['size']-41)<.001,'named actual notes font/print scale'
                text=''.join(chr(x[0])for x in t['chars']);named_text=c['text']
                if c['script']=='Arab':named_text='السَّلامُ'if face=='GeezaPro'else('اﻟﺴ(ﻼَمُ'if expected['variant']=='A'else'اﻟﺴ&ﻼَمُ')
                if c['script']=='Hebr':named_text='שָׁלוֹם'if face=='ArialHebrew'else'שָׁלוֹם'
                assert text==named_text,'named ToUnicode observation'
                assert all(x[1]!=0 for trace in traces for x in trace['chars']),'missing glyph'
                table.mixed.font_resources(doc,page)
                pix=page.get_pixmap(alpha=False);image=Image.frombytes('RGB',(pix.width,pix.height),pix.samples);visible=[]
                for trace in traces:
                    b=trace['bbox'];bounds=(max(0,math.floor(b[0])),max(0,math.floor(b[1])),min(pix.width,math.ceil(b[2])),min(pix.height,math.ceil(b[3])))
                    pixels=image.crop(bounds).tobytes();count=sum(min(pixels[j:j+3])<200 for j in range(0,len(pixels),3));assert count>5,'extractable but invisible notes trace';visible.append({'font':trace['font'],'bounds':bounds,'darkPixels':count})
                if old:
                    assert traces==old[i].get_texttrace(),'native note glyph/font/geometry changed';assert pix.samples==old[i].get_pixmap(alpha=False).samples,'native notes page raster changed'
                source_selection=expected['projections'][i]['selected'];declared=source_selection[0]['family'];match=fonts.get(declared)==face
                rows.append({'id':c['id'],'actualFont':face,'declaredSelection':declared,'declaredSelectionMatchesNative':match,'sourceText':c['text'],'pdfTraceText':text,'size':t['size'],'glyphs':[{'unicode':x[0],'glyph':x[1],'origin':x[2]}for x in t['chars']],'visibleTraceInk':visible,'rasterSha256':hashlib.sha256(pix.samples).hexdigest()})
        finally:
            if old:old.close()
    return{'sha256':sha(file),'controls':rows,'matches':sum(r['declaredSelectionMatchesNative']for r in rows),'differences':sum(not r['declaredSelectionMatchesNative']for r in rows),'fullRasterAndTraceMatchesBaseline':bool(baseline)}

def inspect(directory):
    expected=json.loads((directory/'expected.json').read_text());source=inspect_package(directory/'notes-fonts.pptx',directory,expected,True);pdf=inspect_pdf(directory/'source-notes.pdf',directory,expected);rounds=[]
    for number in [1,2,3]:
        name='native-round'+str(number);package=inspect_package(directory/(name+'.pptx'),directory,expected)
        if rounds:assert package['controls']==rounds[0]['package']['controls']and package['semanticHashes']==rounds[0]['package']['semanticHashes'],'native source domain drift'
        rounds.append({'name':name,'package':package,'pdf':inspect_pdf(directory/(name+'-notes.pdf'),directory,expected,directory/'source-notes.pdf')})
    with fitz.open(directory/'source-notes.pdf')as doc:
        sheet=Image.new('RGB',(595*3,280*5),'white')
        for i,page in enumerate(doc):
            pix=page.get_pixmap(alpha=False);image=Image.frombytes('RGB',(pix.width,pix.height),pix.samples);sheet.paste(image.crop((0,300,595,580)),((i%3)*595,(i//3)*280))
        sheet.save(directory/'notes-contact-sheet.png')
    result={'schemaVersion':1,'directory':str(directory),'sourceAssembly':expected['sourceAssembly'],'source':source,'pdf':pdf,'rounds':rounds,'nativeEdits':'not tested','notesPageLimit':'Slide-image placeholders remain blank in this specimen; complete notes page artwork/fields not certified. Notes PDFs are actual print Notes output, not slide export.'}
    (directory/'notes-inspection.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n');return result

if __name__=='__main__':
    result=inspect(Path(sys.argv[1]));print(json.dumps({'controls':len(result['source']['controls']),'matches':result['pdf']['matches'],'differences':result['pdf']['differences'],'nativeRounds':3}))
