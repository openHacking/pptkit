"""Named Office theme experiment, verified against an earlier native Office asset set."""
import copy, datetime, hashlib, importlib.util, json, math, re, sys, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz
from PIL import Image

spec = importlib.util.spec_from_file_location('table_styles', Path(__file__).with_name('native-table-font-styles.py'))
styles = importlib.util.module_from_spec(spec); spec.loader.exec_module(styles)
append = styles.append; NS = append.NS
ROOT = Path(__file__).resolve().parents[3]
REFERENCE = ROOT / '.tmp/real-font-fresh/run-81jnGm/fresh-theme-sc-round3.pptx'
REFERENCE_SHA = 'a2c196a9007d9ecf64bbab52770414a9d7b054262e6dc5bc9ce3e612a308dc67'
THEME_SHA = '1f98605413696c4fe4029a7b7df696261233d73d8fd06413d3990abfe51ae216'
A16 = 'http://schemas.microsoft.com/office/drawing/2014/main'
GUID = re.compile(r'\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}')
ADDITIONS = {f'ppt/slideLayouts/slideLayout{i}.xml' for i in range(2,12)} | {f'ppt/slideLayouts/_rels/slideLayout{i}.xml.rels' for i in range(2,12)}


def digest_tree(node): return hashlib.sha256(json.dumps(append.tree(node), ensure_ascii=False).encode()).hexdigest()


def office_assets(z):
    hashes, ids, creation, numbers, fields = {}, {}, [], [], {}
    for part in z.namelist():
        if not part.startswith(('ppt/slideMasters/', 'ppt/slideLayouts/')) or not part.endswith(('.xml','.rels')): continue
        node = ET.fromstring(z.read(part)); values = []
        for index, child in enumerate(node.iter()):
            if child.tag in ['{' + A16 + '}creationId', '{' + NS['a'] + '}fld']:
                value = child.attrib['id']; assert GUID.fullmatch(value), 'malformed Office GUID'
                values.append({'node': index, 'tag': child.tag, 'id': value})
                if child.tag.endswith('}fld'): fields.setdefault(child.attrib['type'], set()).add(value)
                else: creation.append(value)
                child.set('id', 'VALIDATED_GUID')
            if child.tag == '{' + append.P14 + '}creationId':
                value = child.attrib['val']; assert value.isdigit() and 0 <= int(value) <= 2**32-1
                values.append({'node': index, 'tag': child.tag, 'val': value}); numbers.append(value); child.set('val', 'VALIDATED_UINT')
        hashes[part] = digest_tree(node); ids[part] = values
    assert len(creation) == len(set(creation)) == 63 and len(numbers) == len(set(numbers)) == 12, 'Office creation identity inventory'
    assert set(fields) == {'datetimeFigureOut','slidenum'} and all(len(v)==1 for v in fields.values()) and len(set.union(*fields.values())) == 2, 'Office shared field identities'
    return {'canonicalHashes': hashes, 'identities': ids, 'sharedFields': {k:next(iter(v)) for k,v in fields.items()}}


def verify_baseline(directory, expected):
    original = Path(expected['sourceNative']['file'])
    assert append.sha(original) == expected['sourceNative']['sha256']
    prior_manifest = json.loads((ROOT / 'docs/testing/native-table-font-style-evidence.json').read_text())
    group = next(g for g in prior_manifest['groups'] if g['group'] == expected['group'].removesuffix('-local'))
    assert group['independent']['stages'][2]['rounds'][2]['package']['sha256'] == expected['sourceNative']['sha256']
    assert append.sha(directory / 'baseline.pptx') == expected['themeBaseline']['pptx']
    styles.inspect_package(original, Path(expected['sourceNative']['directory']), expected, 'regular')
    with zipfile.ZipFile(original) as source, zipfile.ZipFile(directory / 'baseline.pptx') as prepared:
        assert prepared.namelist() == source.namelist()
        for part in source.namelist():
            before, after = source.read(part), prepared.read(part)
            selected = next((i for i in expected['localOverrideTargets'] if part == f'ppt/slides/slide{i}.xml'), None)
            if selected is None: assert before == after, 'baseline patch leaked: ' + part; continue
            left, right = ET.fromstring(before), ET.fromstring(after)
            runs = left.findall('.//a:tbl/a:tr/a:tc/a:txBody/a:p/a:r/a:rPr', NS); assert len(runs)==2
            for run in runs:
                assert len(run)==0
                for slot, family in expected['localFonts'].items(): ET.SubElement(run,'{'+NS['a']+'}'+slot,{'typeface':family})
            assert append.tree(left)==append.tree(right), 'baseline local override patch differs'
    if expected['themeBaseline']['pdf'] is not None:
        assert append.sha(directory / 'baseline.pdf') == expected['themeBaseline']['pdf'] == expected['sourceNative']['pdfSha256']
    assert append.sha(REFERENCE) == REFERENCE_SHA
    prior = json.loads((ROOT / 'docs/testing/real-font-edit-evidence.json').read_text())
    assert prior['results']['theme']['rounds'][2]['package']['sha256'] == REFERENCE_SHA


def read_tables(z, expected, office):
    presentation = ET.fromstring(z.read('ppt/presentation.xml'))
    slides = presentation.findall('p:sldIdLst/p:sldId', NS); assert len(slides)==len(expected['controls'])
    style_part = append.baseline.mixed.edit.relationship(z, 'ppt/presentation.xml', 'tableStyles')
    style_nodes = {s.attrib['styleId']:s for s in ET.fromstring(z.read(style_part)).findall('a:tblStyle',NS)}
    rows=[]
    for index,(slide,control) in enumerate(zip(slides,expected['controls']),1):
        part=append.baseline.mixed.edit.relationship(z,'ppt/presentation.xml','slide',slide.attrib['{'+NS['r']+'}id'])
        root=ET.fromstring(z.read(part)); labels=[s for s in root.findall('p:cSld/p:spTree/p:sp',NS) if s.find('p:nvSpPr/p:cNvPr',NS).attrib['name']=='Label']
        assert len(labels)==1 and ''.join(labels[0].find('p:txBody',NS).itertext())==control['id']
        frames=[f for f in root.findall('p:cSld/p:spTree/p:graphicFrame',NS) if f.find('p:nvGraphicFramePr/p:cNvPr',NS).attrib['name']=='ProbeTable']; assert len(frames)==1
        frame=frames[0]; table=frame.find('a:graphic/a:graphicData/a:tbl',NS);assert len(table.findall('a:tr/a:tc',NS))==1
        style_id=table.find('a:tblPr/a:tableStyleId',NS).text; assert style_id==control['styleId']
        text=style_nodes[style_id].find('a:wholeTbl/a:tcTxStyle',NS)
        layout=append.baseline.mixed.edit.relationship(z,part,'slideLayout'); master=append.baseline.mixed.edit.relationship(z,layout,'slideMaster'); theme=append.baseline.mixed.edit.relationship(z,master,'theme')
        assert layout==('ppt/slideLayouts/slideLayout7.xml' if office else 'ppt/slideLayouts/slideLayout1.xml') and master=='ppt/slideMasters/slideMaster1.xml' and theme=='ppt/theme/theme1.xml'
        theme_root=ET.fromstring(z.read(theme)); ref=text.find('a:fontRef',NS)
        if ref is not None:
            assert ref.attrib['idx']==control['kind']; selected=theme_root.find('a:themeElements/a:fontScheme/a:'+control['kind']+'Font',NS); selector={'type':'reference','collection':control['kind'],'sourcePart':theme}
        else:
            assert control['kind'].startswith('collection'); selected=text.find('a:font',NS); selector={'type':'explicit-collection','sourcePart':style_part}
        runs=table.findall('a:tr/a:tc/a:txBody/a:p/a:r',NS); assert len(runs)==(2 if index in expected['editTargets'] else 1)
        actual=[]
        for number,run in enumerate(runs):
            props=run.find('a:rPr',NS); assert run.find('a:t',NS).text==(control['text'] if number==0 else 'Ω')
            assert props.attrib['sz']=='4800' and props.attrib['lang']==(control['language'] if number==0 else 'el-GR')
            local={slot:props.find('a:'+slot,NS).attrib['typeface'] for slot in ['latin','ea','cs'] if props.find('a:'+slot,NS) is not None}
            assert local==(expected['localFonts'] if index in expected['localOverrideTargets'] else {})
            actual.append({'text':run.find('a:t',NS).text,'attributes':dict(props.attrib),'localFonts':local})
        rows.append({'id':control['id'],'slide':index,'part':part,'layout':layout,'master':master,'theme':theme,'themeName':theme_root.attrib['name'],'selector':selector,'declared':append.baseline.collection(selected),'runs':actual,'localOverride':index in expected['localOverrideTargets']})
    return rows


def inspect_package(file,directory,expected):
    verify_baseline(directory,expected)
    with zipfile.ZipFile(REFERENCE) as reference,zipfile.ZipFile(directory/'baseline.pptx') as before,zipfile.ZipFile(file) as candidate:
        assert len(candidate.namelist())==len(set(candidate.namelist()))
        assert set(candidate.namelist())==set(before.namelist())|ADDITIONS, 'theme part inventory'
        assets=office_assets(candidate); known=office_assets(reference)
        assert assets['canonicalHashes']==known['canonicalHashes'], 'Office assets differ outside validated creation/field identities'
        assert hashlib.sha256(candidate.read('ppt/theme/theme1.xml')).hexdigest()==THEME_SHA==hashlib.sha256(reference.read('ppt/theme/theme1.xml')).hexdigest()
        metadata=[]
        for part in before.namelist():
            old,new=before.read(part),candidate.read(part)
            if part in assets['canonicalHashes'] or part=='ppt/theme/theme1.xml':continue
            if part=='[Content_Types].xml':
                left,right=ET.fromstring(old),ET.fromstring(new)
                items={n.attrib['PartName']:n.attrib['ContentType'] for n in left if n.tag.endswith('}Override')}
                for p in ADDITIONS:
                    if p.endswith('.xml'):items['/'+p]='application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
                actual={n.attrib['PartName']:n.attrib['ContentType'] for n in right if n.tag.endswith('}Override')}
                assert len(actual)==len([n for n in right if n.tag.endswith('}Override')]), 'duplicate content type override'
                assert actual==items and [append.tree(n) for n in left if n.tag.endswith('}Default')]==[append.tree(n) for n in right if n.tag.endswith('}Default')]
            elif part=='ppt/presentation.xml':
                left,right=ET.fromstring(old),ET.fromstring(new); template=ET.fromstring(reference.read(part))
                left.find('p:sldMasterIdLst/p:sldMasterId',NS).set('id',template.find('p:sldMasterIdLst/p:sldMasterId',NS).attrib['id'])
                original=left.find('p:defaultTextStyle',NS);position=list(left).index(original);left.remove(original);left.insert(position,copy.deepcopy(template.find('p:defaultTextStyle',NS)))
                assert append.tree(left)==append.tree(right), 'presentation changes beyond master id/default text style'
            elif part.startswith('ppt/slides/_rels/'):
                left,right=ET.fromstring(old),ET.fromstring(new);found=[]
                for rel in left:
                    if rel.attrib['Type'].endswith('/slideLayout'):found.append(rel); assert rel.attrib['Target']=='../slideLayouts/slideLayout1.xml';rel.set('Target','../slideLayouts/slideLayout7.xml')
                assert len(found)==1 and append.tree(left)==append.tree(right), 'slide layout binding changed unexpectedly'
            elif part=='docProps/app.xml':
                left,right=ET.fromstring(old),ET.fromstring(new);ns=left.tag.split('}')[0]+'}';v='{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}'
                assert left.find(ns+'Template') is None
                left.insert(0,ET.Element(ns+'Template'))
                original=left.find(ns+'TotalTime');actual=right.find(ns+'TotalTime');assert int(actual.text)>=int(original.text);metadata.append({'field':'TotalTime','before':original.text,'after':actual.text});original.text=actual.text
                left.find(ns+'HeadingPairs').find(v+'vector')[1].find(v+'i4').text='3'
                vector=left.find(ns+'TitlesOfParts').find(v+'vector');assert vector[0].text=='Arial' and vector[1].text=='PPTKit'
                titles=[copy.deepcopy(n) for n in list(vector)[2:]];vector.clear();vector.attrib.update({'size':str(len(titles)+4),'baseType':'lpstr'})
                for text in ['等线','等线 Light','Arial','Office 主题\u200b\u200b']:ET.SubElement(vector,v+'lpstr').text=text
                vector.extend(titles);assert append.tree(left)==append.tree(right), 'unexpected app metadata'
            elif part=='docProps/core.xml':
                left,right=ET.fromstring(old),ET.fromstring(new)
                for a,b in zip(left,right):
                    field=a.tag.split('}')[-1]
                    if field=='revision':assert int(b.text)==int(a.text)+1
                    elif field=='modified':assert datetime.datetime.fromisoformat(b.text.replace('Z','+00:00'))>=datetime.datetime.fromisoformat(a.text.replace('Z','+00:00'))
                    else:continue
                    metadata.append({'field':field,'before':a.text,'after':b.text});a.text=b.text
                assert append.tree(left)==append.tree(right)
            else:assert old==new, 'unrelated source part changed: '+part
        rows=read_tables(candidate,expected,True)
        baseline_rows=read_tables(before,expected,False)
        for row,old in zip(rows,baseline_rows):
            assert row['runs']==old['runs']
            if row['selector']['type']=='explicit-collection':assert row['declared']==old['declared']
            else:assert row['declared']!=old['declared']
    return {'sha256':append.sha(file),'controls':rows,'officeAssets':assets,'metadata':metadata,'preservedSlideStyleNotesAndResourceBytes':True}


def body_face(control,expected,office,local):
    if local:return 'ArialUnicodeMS'
    if office and control['kind'] in ['major','minor']:
        role=control['kind'];return {'latin':'DengXian-Light' if role=='major' else 'DengXian-Regular','sc':'DengXian-Light' if role=='major' else 'DengXian-Regular','tc':'PMingLiU','ja':'YuGothic-Light' if role=='major' else 'YuGothic-Regular','ko':'Batang','ko-kr':'MalgunGothic','ar':'TimesNewRomanPSMT' if role=='major' else 'ArialMT','he':'TimesNewRomanPSMT' if role=='major' else 'ArialMT','greek':'DengXian-Light' if role=='major' else 'DengXian-Regular'}[control['sampleId']]
    return append.baseline.observed_face(control,{f['family']:f['postscriptName'] for f in expected['fonts']})


def ink(image,trace):
    b=trace['bbox'];bounds=(max(0,math.floor(b[0])),max(0,math.floor(b[1])),min(720,math.ceil(b[2])),min(260,math.ceil(b[3])))
    pixels=image.crop(bounds).tobytes();count=sum(min(pixels[i:i+3])<200 for i in range(0,len(pixels),3));assert count>5,'invisible glyph trace'
    return {'font':trace['font'],'text':''.join(chr(c[0]) for c in trace['chars']),'bounds':bounds,'darkPixels':count}


def glyphs(traces):return [(t['font'],c) for t in traces for c in t['chars']]


def drawings(page):return [{k:v for k,v in d.items() if k!='seqno'} for d in page.get_drawings()]


def inspect_pdf(file,directory,expected,office=True,previous=None):
    rows=[]
    with fitz.open(file) as doc,fitz.open(directory/'baseline.pdf') as baseline:
        prior=fitz.open(previous) if previous else None
        native=fitz.open(Path(expected['sourceNative']['directory'])/'regular-round3.pdf')
        try:
            assert len(doc)==len(baseline)==len(expected['controls'])
            for index,(page,old,control) in enumerate(zip(doc,baseline,expected['controls']),1):
                assert page.rect==old.rect and abs(page.rect.width-720)<.1 and abs(page.rect.height-260)<.1
                append.baseline.mixed.font_resources(doc,page)
                label=[t for t in page.get_texttrace() if t['bbox'][1]<60]
                assert ''.join(chr(c[0]) for _,c in glyphs(label))==control['id'] and all(f=='ArialMT' and c[1]!=0 for f,c in glyphs(label)), 'label/font/order'
                old_label=[t for t in old.get_texttrace() if t['bbox'][1]<60]
                assert len(glyphs(label))==len(glyphs(old_label))
                for (f,c),(of,oc) in zip(glyphs(label),glyphs(old_label)):
                    assert (f,c[0],c[1],c[2][0])==(of,oc[0],oc[1],oc[2][0]) and abs(c[2][1]-oc[2][1]-(-1.2 if office else 0))<.0001, 'named label baseline response'
                traces=append.body_traces(page);before=append.body_traces(old);target=index in expected['editTargets'];local=index in expected['localOverrideTargets']
                face=body_face(control,expected,office,local)
                greek='ArialMT' if local else ('DengXian-Light' if control['kind']=='major' else 'DengXian-Regular') if office and control['kind'] in ['major','minor'] else 'Georgia'
                assert len(traces)==(2 if target else 1) and all(t['type']==0 and abs(t['size']-48)<.001 for t in traces)
                if target:
                    assert traces[0]['font']==face and traces[1]['font']==greek and ''.join(chr(c[0])for c in traces[1]['chars'])=='Ω', 'named body/Greek face'
                else:assert traces[0]['font']==face
                assert all(c[1]!=0 for _,c in glyphs(traces)), 'missing body glyph'
                actual_text=''.join(chr(c[0])for _,c in glyphs(traces))
                if office and control['kind'] in ['minor','major'] and not local:
                    text={'sc':'⻣直令返','ja':'⾻直令返','ar':'اﻟﺴﱠﻼَمُ' if control['kind']=='major' else 'اﻟﺳﱠﻼَمُ','he':'שָׁלוֹם'}.get(control['sampleId'],control['text'])+('Ω' if target else '')
                    assert actual_text==text, 'named Office ToUnicode observation'
                elif local:
                    text={'ar':'اﻟﺴ*ﻼَمُ','he':'שָׁלוֹם'}.get(control['sampleId'],control['text'])+'Ω';assert actual_text==text
                else:
                    reference=append.body_traces(native[index-1]);assert actual_text==''.join(chr(c[0])for _,c in glyphs(reference)), 'explicit/baseline ToUnicode'
                preserved=office and (control['kind'].startswith('collection') or local)
                delta=-1.2 if office and local and control['kind'] in ['major','minor'] else 0
                if preserved:
                    assert len(glyphs(traces))==len(glyphs(before))
                    for (f,c),(of,oc) in zip(glyphs(traces),glyphs(before)):
                        # Named subset renumbering: retain raw GIDs in evidence; raster is checked separately.
                        gid_delta = -3 if expected['group'] == 'reference' and control['id'] in ['collection-ar','collection-he'] and f == 'ArialMT' else -3 if expected['group'] == 'reference-local' and local and f == 'ArialMT' else 0
                        assert (f,c[0],c[1],c[2][0])==(of,oc[0],oc[1]+gid_delta,oc[2][0]) and abs(c[2][1]-oc[2][1]-delta)<.0001, 'local/collection glyph or position changed beyond named baseline/subset response'
                if local and not office:assert all(abs(c[2][1]-139.68)<.001 for _,c in glyphs(traces)), 'local baseline placement'
                pix=page.get_pixmap(alpha=False);image=Image.frombytes('RGB',(pix.width,pix.height),pix.samples)
                visible=[ink(image,t) for t in label+traces if any(chr(c[0]).strip() for c in t['chars'])]
                if office:
                    assert drawings(page)==drawings(old), 'background/border paths changed'
                    if control['kind'].startswith('collection'):
                        before_pix=old.get_pixmap(alpha=False);before_image=Image.frombytes('RGB',(before_pix.width,before_pix.height),before_pix.samples)
                        assert image.crop((0,60,720,260)).tobytes()==before_image.crop((0,60,720,260)).tobytes(), 'explicit collection body raster changed'
                    assert pix.samples!=old.get_pixmap(alpha=False).samples, 'named theme operation did not change page raster'
                if prior:assert pix.samples==prior[index-1].get_pixmap(alpha=False).samples,'native theme round raster changed'
                rows.append({'id':control['id'],'localOverride':local,'sourceText':control['text']+('Ω' if target else ''),'pdfTraceText':actual_text,'actualFonts':sorted(set(t['font'] for t in traces)),'glyphs':[{'font':f,'unicode':c[0],'glyph':c[1],'origin':c[2]} for f,c in glyphs(traces)],'visibleTraceInk':visible,'preservedOriginalFontUnicodeAndX':preserved,'baselineGlyphs':[{'font':f,'unicode':c[0],'glyph':c[1],'origin':c[2]} for f,c in glyphs(before)] if preserved else None,'namedBodyDeltaY':delta if preserved else None,'labelDeltaY':-1.2 if office else 0,'explicitCollectionBodyRasterPreserved':office and control['kind'].startswith('collection'),'rasterSha256':hashlib.sha256(pix.samples).hexdigest()})
        finally:
            native.close()
            if prior:prior.close()
    return {'sha256':append.sha(file),'controls':rows,'nativeRoundRasterMatch':bool(previous)}


def inspect(directory):
    expected=json.loads((directory/'expected.json').read_text());verify_baseline(directory,expected)
    baseline_pdf=inspect_pdf(directory/'baseline.pdf',directory,expected,False)
    rounds=[]
    for number in range(1,4):
        stem='office-round'+str(number);package=inspect_package(directory/(stem+'.pptx'),directory,expected)
        if rounds:
            assert package['controls']==rounds[0]['package']['controls'] and package['officeAssets']==rounds[0]['package']['officeAssets'],'native theme source/identity drift'
        pdf=inspect_pdf(directory/(stem+'.pdf'),directory,expected,True,directory/'office-round1.pdf' if number>1 else None)
        rounds.append({'name':stem,'package':package,'pdf':pdf})
    with fitz.open(directory/'office-round1.pdf')as doc:
        sheet=Image.new('RGB',(720*3,260*math.ceil(len(doc)/3)),'white')
        for i,page in enumerate(doc):
            pix=page.get_pixmap(alpha=False);sheet.paste(Image.frombytes('RGB',(pix.width,pix.height),pix.samples),((i%3)*720,(i//3)*260))
        sheet.save(directory/'theme-contact-sheet.png')
    result={'schemaVersion':1,'directory':str(directory),'sourceAssembly':expected['sourceAssembly'],'priorOfficeReference':{'file':str(REFERENCE),'sha256':REFERENCE_SHA,'themeSha256':THEME_SHA},'baselinePDF':baseline_pdf,'rounds':rounds,'scope':'Named Office theme/source/native font and placement observations, not universal font resolution or complete typography parity.'}
    (directory/'theme-inspection.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n');return result


if __name__=='__main__':
    result=inspect(Path(sys.argv[1]));print(json.dumps({'controls':len(result['baselinePDF']['controls']),'nativeRounds':3,'addedLayouts':10,'officeReferenceVerified':True}))
