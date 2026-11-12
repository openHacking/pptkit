"""Independent named table fontRef/collection source declarations and native observations."""
import copy,hashlib,importlib.util,json,math,sys,zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import fitz
from PIL import Image
spec=importlib.util.spec_from_file_location('mixed',Path(__file__).with_name('real-font-mixed.py'))
mixed=importlib.util.module_from_spec(spec);spec.loader.exec_module(mixed)
NS=mixed.edit.NS

def collection(node):
    slots=[]
    for key in ['latin','ea','cs']:
        child=node.find('a:'+key,NS);assert child is not None
        slots.append(child.attrib['typeface'])
    rules={}
    for rule in node.findall('a:font',NS):
        assert rule.attrib['script'] not in rules
        rules[rule.attrib['script']]=rule.attrib['typeface']
    return {'slots':slots,'rules':rules}

def inspect_package(file,expected,source=False):
    if source:assert mixed.edit.fresh.sha(file)==expected['sourceSha256'], 'source bytes changed'
    rows=[]
    with zipfile.ZipFile(file) as z:
        assert len(z.namelist())==len(set(z.namelist()))
        presentation=ET.fromstring(z.read('ppt/presentation.xml'))
        styles_part=mixed.edit.relationship(z,'ppt/presentation.xml','tableStyles')
        styles_root=ET.fromstring(z.read(styles_part));styles={}
        for node in styles_root.findall('a:tblStyle',NS):
            assert node.attrib['styleId'] not in styles;styles[node.attrib['styleId']]=node
        slides=presentation.findall('p:sldIdLst/p:sldId',NS);assert len(slides)==len(expected['controls'])
        for slide,control in zip(slides,expected['controls']):
            part=mixed.edit.relationship(z,'ppt/presentation.xml','slide',slide.attrib['{'+NS['r']+'}id'])
            root=ET.fromstring(z.read(part));labels=[n for n in root.findall('p:cSld/p:spTree/p:sp',NS) if n.find('p:nvSpPr/p:cNvPr',NS).attrib['name']=='Label']
            assert len(labels)==1 and ''.join(labels[0].find('p:txBody',NS).itertext())==control['id'], 'source order/label'
            frames=[n for n in root.findall('p:cSld/p:spTree/p:graphicFrame',NS) if n.find('p:nvGraphicFramePr/p:cNvPr',NS).attrib['name']=='ProbeTable']
            assert len(frames)==1;frame=frames[0];table=frame.find('a:graphic/a:graphicData/a:tbl',NS);assert table is not None
            assert len(table.findall('a:tr',NS))==1 and len(table.findall('a:tr/a:tc',NS))==1
            cell=table.find('a:tr/a:tc',NS);runs=cell.findall('a:txBody/a:p/a:r',NS);assert len(runs)==1
            run=runs[0];rpr=run.find('a:rPr',NS)
            assert run.find('a:t',NS).text==control['text'] and rpr.attrib['lang']==control['language'] and rpr.attrib['sz']=='4800', 'text/lang/size'
            assert all(rpr.find('a:'+slot,NS) is None for slot in ['latin','ea','cs']), 'local font masks table selector'
            style_id=table.find('a:tblPr/a:tableStyleId',NS).text;assert style_id==control['styleId'], 'style binding'
            style=styles[style_id];text=style.find('a:wholeTbl/a:tcTxStyle',NS);assert text is not None
            layout=mixed.edit.relationship(z,part,'slideLayout');master=mixed.edit.relationship(z,layout,'slideMaster');theme_part=mixed.edit.relationship(z,master,'theme');theme=ET.fromstring(z.read(theme_part))
            ref=text.find('a:fontRef',NS);explicit=text.find('a:font',NS)
            assert (ref is None)!=(explicit is None)
            if ref is not None:
                assert ref.attrib['idx']==control['kind']
                selected=theme.find('a:themeElements/a:fontScheme/a:'+ref.attrib['idx']+'Font',NS)
                selector={'type':'reference','collection':ref.attrib['idx'],'sourcePart':theme_part}
            else:
                assert control['kind'].startswith('collection');selected=explicit
                selector={'type':'explicit-collection','sourcePart':styles_part}
            declared=collection(selected)
            assert declared['rules']==expected['maps'][control['kind']], 'declared script rules'
            generic='Arial Unicode MS' if control['kind']=='collection-generic' else ''
            latin='Arial' if control['kind']=='minor' else 'Georgia'
            assert declared['slots']==[latin,generic,generic], 'generic source slots'
            assert declared['rules'][control['script']]==control['targetFamily']
            rows.append({'id':control['id'],'part':part,'styleId':style_id,'selector':selector,'declared':declared,'selectedDeclaredFamily':declared['rules'][control['script']],'runText':control['text'],'runAttributes':dict(rpr.attrib),'runXml':ET.tostring(run,encoding='unicode'),'bodyProperties':dict(cell.find('a:txBody/a:bodyPr',NS).attrib),'paragraphProperties':[ET.tostring(p.find('a:pPr',NS),encoding='unicode') if p.find('a:pPr',NS) is not None else None for p in cell.findall('a:txBody/a:p',NS)],'cellProperties':ET.tostring(cell.find('a:tcPr',NS),encoding='unicode'),'geometry':ET.tostring(frame.find('p:xfrm',NS),encoding='unicode'),'styleFontXml':ET.tostring(selected,encoding='unicode')})
    return {'sha256':mixed.edit.fresh.sha(file),'controls':rows}

def expected_native_records(source_rows):
    rows=copy.deepcopy(source_rows)
    for row in rows:
        assert row['bodyProperties']=={'lIns':'88900','tIns':'44450','rIns':'88900','bIns':'44450','anchor':'t'}
        row['bodyProperties']={}
        paragraph=ET.fromstring(row['paragraphProperties'][0])
        assert paragraph.attrib=={'algn':'l','marL':'0','indent':'-0'}
        paragraph.attrib.clear();paragraph.attrib.update({'marL':'0','indent':'0','algn':'l'})
        row['paragraphProperties']=[ET.tostring(paragraph,encoding='unicode')]
        cell=ET.fromstring(row['cellProperties'])
        assert len(cell)==5 and cell[0].tag=='{'+NS['a']+'}noFill'
        for side in ['lnL','lnR','lnT','lnB']:
            line=cell.find('a:'+side,NS)
            assert line is not None and line.attrib=={'w':'12700'} and len(line)==2
            assert line[0].tag=='{'+NS['a']+'}noFill' and line[1].tag=='{'+NS['a']+'}prstDash' and line[1].attrib=={'val':'solid'}
            cell.remove(line)
        row['cellProperties']=ET.tostring(cell,encoding='unicode')
    return rows

def observed_face(control,fonts):
    kind,sample=control['kind'],control['sampleId']
    if kind=='minor' and sample=='ko':return 'MalgunGothic'
    if kind=='major' and sample=='ko':return 'Batang'
    if kind.startswith('collection') and sample not in ['latin','greek']:
        if kind=='collection-generic':return 'ArialUnicodeMS'
        if sample in ['sc','tc','ja']:return 'SimSun'
        if sample in ['ko','ko-kr']:return 'Batang'
        return 'ArialMT'
    return fonts[control['targetFamily']]

def inspect_pdf(file,expected,baseline=None):
    doc=fitz.open(file);assert len(doc)==len(expected['controls'])
    original=fitz.open(baseline) if baseline else None
    fonts={f['family']:f['postscriptName'] for f in expected['fonts']};rows=[];matches=0
    for index,(page,control) in enumerate(zip(doc,expected['controls'])):
        assert abs(page.rect.width-720)<.1 and abs(page.rect.height-260)<.1
        label=''.join(s['text'] for b in page.get_text('dict')['blocks'] if 'lines' in b for line in b['lines'] for s in line['spans'] if s['bbox'][1]<60).strip();assert label==control['id'], 'PDF order/label'
        traces=[t for t in page.get_texttrace() if t['bbox'][1]>60];glyphs=[c for t in traces for c in t['chars']]
        actual=sorted(set(t['font'] for t in traces));observed=observed_face(control,fonts)
        assert actual==[observed], 'named actual PDF face changed'
        assert glyphs and all(c[1]!=0 for c in glyphs), 'missing actual glyph'
        text=''.join(chr(c[0]) for c in glyphs)
        if control['sampleId'] not in ['ar','he']:assert text==control['text'], 'PDF simple-script text changed'
        mixed.font_resources(doc,page)
        pix=page.get_pixmap(alpha=False);image=Image.frombytes('RGB',(pix.width,pix.height),pix.samples);visible=[]
        for trace in traces:
            t=''.join(chr(c[0]) for c in trace['chars'])
            if not t.strip():continue
            b=trace['bbox'];bounds=(max(0,math.floor(b[0])),max(0,math.floor(b[1])),min(720,math.ceil(b[2])),min(260,math.ceil(b[3])))
            ink=sum(min(pixel)<200 for pixel in image.crop(bounds).getdata());assert ink>5, 'font span invisible'
            visible.append({'font':trace['font'],'text':t,'bounds':bounds,'darkPixels':ink})
        if original:
            before=original[index].get_pixmap(alpha=False);assert (pix.width,pix.height,pix.samples)==(before.width,before.height,before.samples), 'native round raster changed'
        matched=observed==fonts[control['targetFamily']];matches+=matched
        rows.append({'id':control['id'],'declaredFamily':control['targetFamily'],'declaredPostscriptName':fonts[control['targetFamily']],'actualFonts':actual,'declaredSelectionMatchesNative':matched,'sourceText':control['text'],'pdfTraceText':text,'glyphs':[{'unicode':c[0],'glyph':c[1],'origin':c[2]} for c in glyphs],'glyphZeroCount':0,'visibleTraceInk':visible,'rasterSha256':hashlib.sha256(pix.samples).hexdigest()})
    assert matches==(18 if len(rows)==27 else 2), 'named match/difference counts changed'
    return {'sha256':mixed.edit.fresh.sha(file),'matches':matches,'differences':len(rows)-matches,'controls':rows,'baselineRasterMatch':bool(original)}

def inspect(directory):
    expected=json.loads((directory/'expected.json').read_text());assert len(expected['controls']) in [27,9]
    for f in expected['fonts']:assert mixed.edit.fresh.sha(f['file'])==f['sha256']
    source=inspect_package(directory/'table-fonts.pptx',expected,True);pdf=inspect_pdf(directory/'table-fonts.pdf',expected);rounds=[]
    for round in range(1,4):
        stem='table-fonts-round'+str(round);package=inspect_package(directory/(stem+'.pptx'),expected)
        assert package['controls']==expected_native_records(source['controls']), 'selected native source records differ from named normalization'
        rounds.append({'name':stem,'package':package,'pdf':inspect_pdf(directory/(stem+'.pdf'),expected,baseline=directory/'table-fonts.pdf')})
    with fitz.open(directory/'table-fonts.pdf') as doc:
        count=len(doc);sheet=Image.new('RGB',(720*3,260*((count+2)//3)),'white')
        for i,page in enumerate(doc):
            pix=page.get_pixmap(alpha=False);sheet.paste(Image.frombytes('RGB',(pix.width,pix.height),pix.samples),((i%3)*720,(i//3)*260))
        sheet.save(directory/'table-fonts-contact-sheet.png')
    result={'schemaVersion':1,'artifactsDirectory':str(directory),'source':source,'pdf':pdf,'rounds':rounds,'nativeEdits':'not tested','sourceSelectionScope':'Declared v3 source maps are retained and checked independently; actual native differences are explicit, not a native-parity claim.'}
    (directory/'table-fonts-inspection.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
    return result

if __name__=='__main__':
    result=inspect(Path(sys.argv[1]));print(json.dumps({'controls':len(result['source']['controls']),'matches':result['pdf']['matches'],'differences':result['pdf']['differences'],'nativeRounds':3}))
