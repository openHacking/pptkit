"""Original author map -> three families -> local native editing, for six named pages."""
import hashlib, importlib.util, json, re, sys
from pathlib import Path
import fitz
from PIL import Image
spec = importlib.util.spec_from_file_location('edit', Path(__file__).with_name('real-font-edit.py'))
edit = importlib.util.module_from_spec(spec); spec.loader.exec_module(edit)
STAGES = {'edit':'mixed-edit', 'bold':'mixed-bold-final', 'italic':'mixed-italic', 'theme':'mixed-theme'}

def wanted_runs(control, stage):
    if stage == 'source':
        return [{'text':r['text'],'language':control['language'],'altLanguage':None,'size':'4800','bold':False,'italic':False,'slots':[r['family']]*3} for r in control['expectedRuns']]
    language = control['language']
    values = [('Latin',language,None,'Arial'), ('X','en',None if control['fontId'] in ['ar','he'] else language,'Arial'), (' ',language,None,'Arial'), (control['nativeText']+' ',language,None,control['expectedRuns'][1]['family']), ('Ω',language,None,'Georgia')]
    return [{'text':text,'language':lang,'altLanguage':alt,'size':'4800','bold':stage=='bold' and i<2,'italic':stage=='italic' and i<2,'slots':[family]*3} for i,(text,lang,alt,family) in enumerate(values)]

def inspect_package(file, expected, source_rows=None, stage='source'):
    rows = edit.read_package(file,6)
    if stage == 'source': assert edit.fresh.sha(file)==expected['sourceSha256'], 'source bytes changed'
    for index,(row,control) in enumerate(zip(rows,expected['controls'])):
        assert row['id']==control['id'], 'order/label'
        assert [edit.semantic_run(r) for r in row['runs']]==wanted_runs(control,stage), 'text/language/font/topology/local style'
        if source_rows is not None:
            source = source_rows[index]
            assert row['geometry']==source['geometry'] and row['bodyProperties']==source['bodyProperties'], 'frame changed'
            assert row['paragraphProperties']==[None], 'native paragraph properties'
            if stage=='theme':
                assert row['theme']['name'].startswith('Office') and row['theme']['sha256']!=source['theme']['sha256'] and row['theme']['fontScheme']!=source['theme']['fontScheme'], 'theme unchanged'
            else:
                # The first native save rewrites theme serialization; the independent prior native baseline pins its bytes.
                assert row['theme']['name']==source['theme']['name'] and row['theme']['fontScheme']==source['theme']['fontScheme'], 'unexpected font-scheme change'
                assert row['theme']['sha256']=='45fd4ae0d4ed96634406c7e5e39afa1102f5a94e5dbdc44cd52889cab01efa49', 'unexpected normalized native theme'
    return {'sha256':edit.fresh.sha(file),'controls':rows}

def font_resources(doc, page):
    resources={f[4]:f[3].split('+')[-1] for f in page.get_fonts()}
    for font in page.get_fonts():
        kind, declared=doc.xref_get_key(font[0],'BaseFont')
        assert kind=='name' and declared.lstrip('/').split('+')[-1]==font[3].split('+')[-1], 'font resource/reported-name conflict'
        descriptor_kind, descriptor_name=doc.xref_get_key(font[0],'FontDescriptor/FontName')
        if descriptor_kind=='null':
            descendant_kind, descendants=doc.xref_get_key(font[0],'DescendantFonts')
            refs=re.findall(r'(\d+)\s+0\s+R',descendants)
            assert descendant_kind=='array' and len(refs)==1, 'unexpected named font descendants'
            descriptor_kind, descriptor_name=doc.xref_get_key(int(refs[0]),'FontDescriptor/FontName')
        assert descriptor_kind=='name' and descriptor_name.lstrip('/').split('+')[-1]==declared.lstrip('/').split('+')[-1], 'font resource/descriptor-name conflict'
    return resources

def inspect_pdf(file,expected,stage,baseline=None):
    doc=fitz.open(file);assert len(doc)==6
    previous=fitz.open(baseline) if baseline else None
    fonts={f['id']:f['postscriptName'] for f in expected['fonts']};rows=[]
    for index,(page,control) in enumerate(zip(doc,expected['controls'])):
        assert abs(page.rect.width-720)<.1 and abs(page.rect.height-260)<.1
        label=''.join(s['text'] for b in page.get_text('dict')['blocks'] if 'lines' in b for line in b['lines'] for s in line['spans'] if s['bbox'][1]<60).strip()
        assert label==control['id'], 'PDF order/label'
        traces=[t for t in page.get_texttrace() if t['bbox'][1]>60];glyphs=[c for t in traces for c in t['chars']]
        assert glyphs and all(c[1]!=0 for c in glyphs), 'missing glyph'
        latin='Arial-BoldMT' if stage=='bold' else 'Arial-ItalicMT' if stage=='italic' else 'ArialMT'
        native=fonts[control['fontId']];actual=sorted(set(t['font'] for t in traces))
        assert actual==sorted([latin,native,'Georgia']), 'actual font/local style bleed'
        for trace in traces:
            for char in trace['chars']:
                if char[0]==32: continue  # Source spaces are checked exactly in XML; PDF may encode only their advances.
                desired='Georgia' if char[0]==937 else latin if chr(char[0]) in 'LatinX' else native
                assert trace['font']==desired, 'actual face per glyph'
        text=''.join(chr(c[0]) for c in glyphs)
        if control['fontId']=='ar': tail='السَّلامُ Ω'
        elif control['fontId']=='he': tail='שָׁלוֹם Ω'
        else: tail=control['nativeText']+'Ω'
        prefix=('Latin ' if control['fontId'] not in ['ar','he'] else 'Latin') if stage=='source' else 'LatinX'
        assert text==prefix+tail, 'named PDF trace changed'
        resources=font_resources(doc,page)
        stream=b''.join(doc.xref_stream(x) for x in page.get_contents()).decode('latin1')
        pattern=r'BT\s+(?:[-\d.]+\s+Tc\s+)?([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+Tm\s+/([A-Za-z0-9]+)\s+1\s+Tf'
        matrices=[]
        for match in re.findall(pattern,stream):
            face=resources[match[6]];matrix=[float(x) for x in match[:6]]
            # ArialMT is shared with the 16pt label; only the 48pt matrices belong to the probe.
            if face not in actual or matrix[0]!=200: continue
            assert matrix[1]==matrix[2]==0 and matrix[3]==200, 'unexpected synthetic shear'
            matrices.append({'face':face,'matrix':matrix})
        assert set(m['face'] for m in matrices)==set(actual), 'missing body matrices'
        pix=page.get_pixmap(alpha=False);image=Image.frombytes('RGB',(pix.width,pix.height),pix.samples)
        ink=sum(min(p)<200 for p in image.crop((0,60,720,260)).getdata());assert ink>100
        if previous:
            old=previous[index].get_pixmap(alpha=False)
            assert (pix.width,pix.height,pix.samples)==(old.width,old.height,old.samples), 'round raster changed'
        source_text=control['text'] if stage=='source' else control['text'].replace('Latin ','LatinX ',1)
        rows.append({'id':control['id'],'fonts':actual,'sourceText':source_text,'pdfTraceText':text,'pdfTraceTextEqualsSource':text==source_text,'glyphCount':len(glyphs),'glyphZeroCount':0,'darkPixels':ink,'bodyMatrices':matrices,'glyphsByFont':[{'font':t['font'],'glyphs':[{'unicode':c[0],'glyph':c[1],'origin':c[2]} for c in t['chars']]} for t in traces],'rasterSha256':hashlib.sha256(pix.samples).hexdigest()})
    return {'sha256':edit.fresh.sha(file),'controls':rows,'baselineRasterMatch':bool(previous)}

def inspect(directory):
    expected=json.loads((directory/'expected.json').read_text())
    for font in expected['fonts']:assert edit.fresh.sha(font['file'])==font['sha256']
    source=inspect_package(directory/'mixed-fonts.pptx',expected)
    results={'source':{'package':source,'pdf':inspect_pdf(directory/'mixed-fonts.pdf',expected,'source')}}
    sheet=Image.new('RGB',(720*5,260*6),'white')
    for col,(stage,stem) in enumerate({'source':'mixed-fonts',**STAGES}.items()):
        pdf=directory/(stem+'.pdf')
        if stage!='source':
            results[stage]={'initial':{'package':inspect_package(directory/(stem+'.pptx'),expected,source['controls'],stage),'pdf':inspect_pdf(pdf,expected,stage)},'rounds':[]}
            for round in range(1,4):
                name=stem+'-round'+str(round)
                package=inspect_package(directory/(name+'.pptx'),expected,source['controls'],stage)
                assert package['controls']==results[stage]['initial']['package']['controls'], 'raw fields changed across round'
                results[stage]['rounds'].append({'name':name,'package':package,'pdf':inspect_pdf(directory/(name+'.pdf'),expected,stage,baseline=pdf)})
        with fitz.open(pdf) as doc:
            for row,page in enumerate(doc):
                pix=page.get_pixmap(alpha=False);sheet.paste(Image.frombytes('RGB',(pix.width,pix.height),pix.samples),(col*720,row*260))
    for i in range(6):
        def neighbors(row):
            return [(t['font'],[(g['unicode'],g['glyph']) for g in t['glyphs']]) for t in row['glyphsByFont'] if t['font'] not in ['ArialMT','Arial-BoldMT','Arial-ItalicMT']]
        original_neighbors=neighbors(results['source']['pdf']['controls'][i])
        for stage in STAGES:
            assert neighbors(results[stage]['initial']['pdf']['controls'][i])==original_neighbors, 'neighbor glyph sequence changed'
        regular=results['edit']['initial']['pdf']['controls'][i]['rasterSha256']
        assert all(results[stage]['initial']['pdf']['controls'][i]['rasterSha256']!=regular for stage in ['bold','italic','theme']), 'operation invisible'
    sheet.save(directory/'mixed-contact-sheet.png')
    return {'schemaVersion':1,'application':'Microsoft PowerPoint 16.104 (16.104.25121423)','artifactsDirectory':str(directory),'sourceSha256':expected['sourceSha256'],'warnings':expected['warnings'],'scope':'six original author mixed-script runs; three materialized families; Latin-only insertion/bold/italic; Office theme; three native reopen rounds per operation','results':results,'pdfToUnicodeCaveat':'Spaces may be advance-only; Arabic fatha and Hebrew presentation-form differences remain explicit. XML text is checked independently. This is not a complete bidi or glyph-outline certification.'}

if __name__=='__main__':
    directory=Path(sys.argv[1]);result=inspect(directory)
    (directory/'mixed-inspection.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps({'slides':6,'sourceRunsPerSlide':3,'editedRunsPerSlide':5,'stages':list(STAGES),'roundsPerStage':3}))
