"""Independent Office/notes-domain evidence, including native renumbering and named header kerning."""
import copy, hashlib, importlib.util, json, math, posixpath, sys, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz
from PIL import Image, ImageDraw

def module(name,file):
    spec=importlib.util.spec_from_file_location(name,Path(__file__).with_name(file));value=importlib.util.module_from_spec(spec);spec.loader.exec_module(value);return value
notes=module('notes','native-notes-fonts.py')
office=module('office','native-table-font-themes.py')
NS,A,P=notes.NS,notes.A,notes.P
ROOT=Path(__file__).resolve().parents[3]
ANCILLARY={'docProps/core.xml','docProps/app.xml','ppt/viewProps.xml'}
REL='http://schemas.openxmlformats.org/package/2006/relationships'

def optional(z,owner,kind):
    slash=owner.rfind('/');part=owner[:slash+1]+'_rels/'+owner[slash+1:]+'.rels'
    if part not in z.namelist():return None
    nodes=[n for n in notes.root(z,part)if n.get('Type').endswith('/'+kind)]
    assert len(nodes)<=1
    return notes.relation(z,owner,kind) if nodes else None

def domains(z,controls):
    presentation=notes.root(z,'ppt/presentation.xml');ids=presentation.findall('p:sldIdLst/p:sldId',NS)
    assert len(ids)==len(controls)
    masters=presentation.findall('p:notesMasterIdLst/p:notesMasterId',NS);assert len(masters)==1
    listed=notes.relation(z,'ppt/presentation.xml','notesMaster',masters[0].get('{'+NS['r']+'}id'))
    actual_masters=[p for p in z.namelist()if p.startswith('ppt/notesMasters/') and p.endswith('.xml')];assert actual_masters==[listed]
    rows=[]
    for index,(slide,c)in enumerate(zip(ids,controls),1):
        part=notes.relation(z,'ppt/presentation.xml','slide',slide.get('{'+NS['r']+'}id'));note=notes.relation(z,part,'notesSlide')
        assert notes.relation(z,note,'slide')==part
        master=notes.relation(z,note,'notesMaster');assert master==listed
        theme=notes.relation(z,master,'theme');override=optional(z,note,'themeOverride')
        layout=notes.relation(z,part,'slideLayout');slide_master=notes.relation(z,layout,'slideMaster');slide_theme=notes.relation(z,slide_master,'theme')
        assert slide_theme!=theme and override!=slide_theme
        label=f"notes-{c['variant']}-{c['id']}"
        assert notes.named(notes.root(z,part),'SlideLabel').find('p:txBody/a:p/a:r/a:t',NS).text==label
        assert notes.named(notes.root(z,note),'Notes Control Label').find('p:txBody/a:p/a:r/a:t',NS).text==label
        assert notes.body(notes.root(z,note)).find('p:nvSpPr/p:nvPr/p:ph',NS).attrib=={'type':'body','idx':'7'}
        assert notes.body(notes.root(z,master)).find('p:nvSpPr/p:nvPr/p:ph',NS).attrib=={'type':'body','idx':'99'}
        rows.append({'page':index,'id':c['id'],'variant':c['variant'],'slideId':slide.get('id'),'slidePart':part,'notesPart':note,'masterPart':master,'masterThemePart':theme,'overridePart':override,'effectiveThemePart':override or theme,'layoutPart':layout,'slideMasterPart':slide_master,'slideThemePart':slide_theme})
    return rows

def baseline_override(file,directory,expected,controls):
    source=directory/'control-distinct-overrides.pptx';assert notes.sha(source)==expected['distinctSourceSha256']
    shared=directory/'source.pptx';assert notes.sha(shared)==expected['sourceSha256']
    with zipfile.ZipFile(shared)as original,zipfile.ZipFile(source)as distinct:
        original_rows,distinct_rows=domains(original,controls),domains(distinct,controls)
        assert len(distinct.namelist())==len(original.namelist())+14
        shared_part='ppt/theme/notes-font-override-B.xml';assert shared_part not in distinct.namelist()
        override_parts={r['overridePart']for r in distinct_rows if r['overridePart']}
        assert len(override_parts)==15 and set(distinct.namelist())==(set(original.namelist())-{shared_part})|override_parts
        for left,right in zip(original_rows,distinct_rows):
            assert {k:v for k,v in left.items()if k not in ['overridePart','effectiveThemePart']}=={k:v for k,v in right.items()if k not in ['overridePart','effectiveThemePart']}
            if left['variant']=='B':assert notes.tree(notes.root(original,shared_part))==notes.tree(notes.root(distinct,right['overridePart']))
            parent,name=left['notesPart'].rsplit('/',1);rel=parent+'/_rels/'+name+'.rels'
            x,y=notes.root(original,rel),notes.root(distinct,rel)
            if left['variant']=='B':
                node=next(n for n in x if n.get('Type').endswith('/themeOverride'))
                node.set('Target','../theme/'+Path(right['overridePart']).name)
            assert notes.tree(x)==notes.tree(y),'distinct conversion changed non-target relationship'
        for part in original.namelist():
            if part==shared_part or part=='[Content_Types].xml' or part.startswith('ppt/notesSlides/_rels/'):continue
            assert original.read(part)==distinct.read(part),'distinct conversion changed non-override payload'
    with zipfile.ZipFile(source)as old,zipfile.ZipFile(file)as z:
        before,after=domains(old,controls),domains(z,controls);mapping={}
        assert len(z.namelist())==len(set(z.namelist()))==len(old.namelist())-1
        for left,right in zip(before,after):
            assert left['slideId']==right['slideId']
            for key in ['slidePart','notesPart','masterPart','masterThemePart','overridePart']:
                if left[key]:mapping[left[key]]=right[key]
            assert bool(right['overridePart'])==(right['variant']=='B')
            x,y=notes.root(old,left['slidePart']),notes.root(z,right['slidePart']);notes.normalize_paragraphs(x);assert notes.tree(x)==notes.tree(y),'native slide renumbering changed contents'
            x,y=notes.root(old,left['notesPart']),notes.root(z,right['notesPart']);notes.normalize_paragraphs(x);notes.add_locks(x)
            if left['id']=='latin':
                p=notes.body(x).find('p:txBody/a:p',NS);end=p.find('a:endParaRPr',NS);assert end.attrib=={'lang':'en-US'};p.remove(end)
            assert notes.tree(x)==notes.tree(y),'native notes normalization exceeds named locks/zero/Latin ending'
            assert notes.tree(notes.root(old,left['effectiveThemePart']))==notes.tree(notes.root(z,right['effectiveThemePart'])),'native notes font scheme lost'
        assert len({r['overridePart']for r in after if r['overridePart']})==15
        x,y=notes.root(old,before[0]['masterPart']),notes.root(z,after[0]['masterPart']);notes.add_locks(x)
        r=x.find('p:notesStyle/a:lvl1pPr/a:defRPr',NS);paint=r.find('a:solidFill',NS);assert paint.find('a:srgbClr',NS).get('val')=='000000';r.remove(paint)
        assert notes.tree(x)==notes.tree(y),'notes master normalization'
        for part in ['ppt/slideMasters/slideMaster1.xml','ppt/slideLayouts/slideLayout1.xml','ppt/theme/theme1.xml','ppt/tableStyles.xml']:
            assert notes.tree(notes.root(old,part))==notes.tree(notes.root(z,part)),'unrelated source tree changed'
        def rel_owner(part):
            prefix,name=part.rsplit('/_rels/',1) if '/_rels/'in part else ('',part.removeprefix('_rels/'))
            return (prefix+'/' if prefix else '')+name.removesuffix('.rels')
        def rel_part(owner):
            parent,name=owner.rsplit('/',1) if '/'in owner else ('',owner)
            return (parent+'/' if parent else '')+'_rels/'+name+'.rels'
        for part in old.namelist():
            if not part.endswith('.rels'):continue
            owner=rel_owner(part);native_owner=mapping.get(owner,owner);native_part=rel_part(native_owner)
            def edges(archive,p,base,rename):
                output=[]
                for n in notes.root(archive,p):
                    assert n.get('TargetMode','Internal')=='Internal'
                    target=posixpath.normpath(posixpath.join(posixpath.dirname(base),n.get('Target')))
                    output.append((n.get('Type'),mapping.get(target,target) if rename else target))
                return sorted(output)
            assert edges(old,part,owner,True)==edges(z,native_part,native_owner,False),'native relationship role changed: '+part
        return {'sha256':notes.sha(file),'domains':after,'sourceToNativeParts':mapping,'notesSemantics':{r['notesPart']:notes.digest(notes.root(z,r['notesPart']))for r in after},'scope':'Selected complete content/master/font trees and all relationship targets; presentation/metadata native normalization not asserted as byte preservation.'}

def inspect_theme_package(file,directory,expected,controls):
    baseline=directory/('native-baseline-round3.pptx' if expected['group']=='override' else 'baseline.pptx')
    with zipfile.ZipFile(baseline)as old,zipfile.ZipFile(file)as z,zipfile.ZipFile(office.REFERENCE)as reference:
        assert notes.sha(office.REFERENCE)==office.REFERENCE_SHA
        assert len(z.namelist())==len(set(z.namelist()))
        assert set(z.namelist())==set(old.namelist())|office.ADDITIONS,'Office part inventory'
        assets=office.office_assets(z);assert assets['canonicalHashes']==office.office_assets(reference)['canonicalHashes'],'unexpected Office assets'
        assert hashlib.sha256(z.read('ppt/theme/theme1.xml')).hexdigest()==office.THEME_SHA
        rows=domains(z,controls);before=domains(old,controls)
        for left,right in zip(before,rows):
            assert {k:v for k,v in left.items()if k!='layoutPart'}=={k:v for k,v in right.items()if k!='layoutPart'}
            assert right['layoutPart']=='ppt/slideLayouts/slideLayout7.xml'
            assert notes.root(z,right['layoutPart']).get('type')=='blank'
        recorded={}
        for part in old.namelist():
            if part in ANCILLARY:
                recorded[part]={'beforeXml':old.read(part).decode(),'afterXml':z.read(part).decode()};continue
            if part.startswith(('ppt/slideMasters/','ppt/slideLayouts/')) or part=='ppt/theme/theme1.xml':continue
            if part=='[Content_Types].xml':
                x,y=notes.root(old,part),notes.root(z,part)
                assert len(y)==len(x)+10
                entries={n.get('PartName'):dict(n.attrib)for n in y if n.tag.endswith('}Override')}
                assert len(entries)==sum(n.tag.endswith('}Override')for n in y)
                for n in x:
                    if n.tag.endswith('}Override'):assert entries[n.get('PartName')]==dict(n.attrib)
                    else:assert any(notes.tree(n)==notes.tree(c)for c in y)
                for i in range(2,12):assert entries[f'/ppt/slideLayouts/slideLayout{i}.xml']['ContentType']=='application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
                continue
            if part=='ppt/presentation.xml':
                x,y=notes.root(old,part),notes.root(z,part)
                master=x.find('p:sldMasterIdLst/p:sldMasterId',NS);assert master.get('id')=='2147483648';master.set('id','2147483650')
                default=x.find('p:defaultTextStyle',NS);position=list(x).index(default);x.remove(default);x.insert(position,copy.deepcopy(notes.root(reference,part).find('p:defaultTextStyle',NS)))
                assert notes.tree(x)==notes.tree(y),'presentation changed beyond master ID/Office default text styles'
                recorded[part]={'beforeXml':old.read(part).decode(),'afterXml':z.read(part).decode(),'allowed':'master ID 2147483648→2147483650 and complete known Office defaultTextStyle'}
                continue
            if part.startswith('ppt/slides/_rels/'):
                assert z.read(part)==old.read(part).replace(b'../slideLayouts/slideLayout1.xml',b'../slideLayouts/slideLayout7.xml'),'slide relation changed beyond blank Office layout'
            else:assert z.read(part)==old.read(part),'unrelated theme payload changed: '+part
        return {'sha256':notes.sha(file),'domains':rows,'officeAssets':assets,'notesPayloadsExact':True,'normalizationRecords':recorded}

def adjusted_header(trace):
    result=copy.deepcopy(trace);chars=list(result['chars']);last=chars[-1];assert last[0]==65
    delta=-0.917999267578125
    origin=(last[2][0]+delta,last[2][1]);b=last[3];box=(b[0]+delta,b[1],b[2]+delta,b[3]);chars[-1]=(last[0],last[1],origin,box);result['chars']=tuple(chars)
    b=result['bbox'];result['bbox']=(b[0],b[1],b[2]+delta,b[3]);return result

def inspect_pdf(file,baseline,controls,theme=False,reference=None):
    records=[]
    with fitz.open(file)as doc,fitz.open(baseline)as source:
        assert len(doc)==len(source)==len(controls)
        prior=fitz.open(reference)if reference else None
        try:
            for i,(page,old,c)in enumerate(zip(doc,source,controls),1):
                assert tuple(page.rect)==tuple(old.rect)==(0,0,595,842)
                traces,before=page.get_texttrace(),old.get_texttrace();expected=copy.deepcopy(before)
                headers=[j for j,t in enumerate(before)if 350<t['bbox'][1]<410];assert len(headers)==1
                shift=theme and ''.join(chr(x[0])for x in before[headers[0]]['chars'])=='Independent Notes A'
                if shift:expected[headers[0]]=adjusted_header(expected[headers[0]])
                assert traces==expected,'notes glyph/font/geometry changed outside named final-header-A kerning'
                assert all(x[1]!=0 for t in traces for x in t['chars'])
                body=[t for t in traces if t['bbox'][1]>410];assert body and body[0]['font']==notes.observed(c,c['variant'])
                assert all(abs(t['size']-41)<.001 for t in body)
                text=c['text'];face=body[0]['font']
                if c['script']=='Arab':
                    # The combined 30-page PDF assigns a different subset ToUnicode
                    # value to A/local-ar than the separate A-only PDF. Preserve it.
                    combined=len({control['variant']for control in controls})==2
                    text='السَّلامُ'if face=='GeezaPro'else('اﻟﺴ(ﻼَمُ'if c['variant']=='A'and not combined else'اﻟﺴ&ﻼَمُ')
                if c['script']=='Hebr':text='שָׁלוֹם'if face=='ArialHebrew'else'שָׁלוֹם'
                assert ''.join(chr(x[0])for x in body[0]['chars'])==text,'named native ToUnicode observation'
                notes.table.mixed.font_resources(doc,page)
                pix,oldpix=page.get_pixmap(alpha=False),old.get_pixmap(alpha=False)
                image=Image.frombytes('RGB',(pix.width,pix.height),pix.samples);oldimage=Image.frombytes('RGB',(oldpix.width,oldpix.height),oldpix.samples)
                visible=[]
                for trace in traces:
                    b=trace['bbox'];bounds=(math.floor(b[0]),math.floor(b[1]),math.ceil(b[2]),math.ceil(b[3]));pixels=image.crop(bounds).tobytes();ink=sum(min(pixels[j:j+3])<200 for j in range(0,len(pixels),3));assert ink>5;visible.append(ink)
                if shift:
                    b=before[headers[0]]['chars'][-1][3];a=traces[headers[0]]['chars'][-1][3];bounds=(math.floor(min(a[0],b[0]))-1,math.floor(a[1])-1,math.ceil(max(a[2],b[2]))+1,math.ceil(a[3])+1)
                    pixels=image.crop((math.floor(a[0]),math.floor(a[1]),math.ceil(a[2]),math.ceil(a[3]))).tobytes()
                    assert sum(min(pixels[j:j+3])<200 for j in range(0,len(pixels),3))>20,'extractable but invisible shifted header A'
                    ImageDraw.Draw(image).rectangle(bounds,fill='white');ImageDraw.Draw(oldimage).rectangle(bounds,fill='white')
                assert image.tobytes()==oldimage.tobytes(),'notes raster changed outside named final-header glyph'
                if prior:
                    assert traces==prior[i-1].get_texttrace();assert pix.samples==prior[i-1].get_pixmap(alpha=False).samples,'later notes round drift'
                records.append({'page':i,'id':c['id'],'variant':c['variant'],'bodyFonts':[t['font']for t in body],'bodyPdfTraceText':[''.join(chr(x[0])for x in t['chars'])for t in body],'bodyTraceExact':True,'fullRasterExact':pix.samples==oldpix.samples,'headerFinalAGlyphShiftX':-0.917999267578125 if shift else 0,'allTraceInk':visible,'rasterSha256':hashlib.sha256(pix.samples).hexdigest()})
        finally:
            if prior:prior.close()
    return {'sha256':notes.sha(file),'pages':records}

def inspect(directory):
    expected=json.loads((directory/'expected.json').read_text());group=expected['group']
    controls=json.loads((directory/'controls.json').read_text())if group=='override'else[{**c,'variant':group}for c in expected['controls']]
    baseline_rounds=[]
    if group=='override':
        for number in [1,2,3]:
            name=f'native-baseline-round{number}';package=baseline_override(directory/(name+'.pptx'),directory,expected,controls)
            if baseline_rounds:assert package=={**baseline_rounds[0]['package'],'sha256':package['sha256']}
            baseline_rounds.append({'name':name,'package':package,'pdf':inspect_pdf(directory/(name+'-notes.pdf'),directory/'source-distinct-notes.pdf',controls)})
    else:
        assert notes.sha(directory/'baseline.pptx')==expected['themeBaseline']['pptx']
        assert notes.sha(directory/'baseline-notes.pdf')==expected['themeBaseline']['pdf']
    rounds=[];baseline=directory/('native-baseline-round3-notes.pdf'if group=='override'else'baseline-notes.pdf')
    for number in [1,2,3]:
        name=f'office-round{number}';package=inspect_theme_package(directory/(name+'.pptx'),directory,expected,controls)
        if rounds:assert package['domains']==rounds[0]['package']['domains']and package['officeAssets']==rounds[0]['package']['officeAssets']
        rounds.append({'name':name,'package':package,'pdf':inspect_pdf(directory/(name+'-notes.pdf'),baseline,controls,True,directory/'office-round1-notes.pdf'if number>1 else None)})
    result={'group':group,'baselineRounds':baseline_rounds,'themeRounds':rounds,'limits':'Named Office slide theme with independent notes scheme/overrides. Known final-header-A kerning remains explicit. Blank slide-image/full page certification and source style writer remain open.'}
    (directory/'theme-inspection.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n');return result

if __name__=='__main__':
    result=inspect(Path(sys.argv[1]));print(json.dumps({'group':result['group'],'themeRounds':3,'baselineRounds':len(result['baselineRounds'])}))
