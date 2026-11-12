"""One actual paragraph: Latin/SC/TC/ja/ko/ar/he plus Greek and native local edits."""
import copy,hashlib,importlib.util,json,math,re,sys
from pathlib import Path
import fitz
from PIL import Image
spec=importlib.util.spec_from_file_location('mixed',Path(__file__).with_name('real-font-mixed.py'))
mixed=importlib.util.module_from_spec(spec);spec.loader.exec_module(mixed)
edit=mixed.edit
STAGES=['edit','bold','italic','theme']

def wanted_runs(expected,stage,native_round=0):
    original=expected['expectedRuns']
    if stage=='source': values=original
    elif stage=='edit' and not native_round:values=[{**original[0],'text':'LatinX '},*original[1:]]
    else: values=[{**original[0],'text':'LatinX'},{**original[0],'text':' '},*original[1:]]
    return [{'text':r['text'],'language':r['language'],'altLanguage':None,'size':'4800','bold':stage=='bold' and i==0,'italic':stage=='italic' and i==0,'slots':[r['family']]*3} for i,r in enumerate(values)]

def inspect_package(file,expected,source=None,stage='source',native_round=0):
    rows=edit.read_package(file,1);row=rows[0]
    assert row['id']=='combined-scripts'
    assert [edit.semantic_run(r) for r in row['runs']]==wanted_runs(expected,stage,native_round), 'source text/lang/topology/local font/style'
    if stage=='source':assert edit.fresh.sha(file)==expected['sourceSha256']
    else:
        assert row['geometry']==source['geometry'] and row['bodyProperties']==source['bodyProperties'], 'frame changed'
        assert row['paragraphProperties']==[None], 'native paragraph properties'
        if stage=='theme':
            assert row['theme']['name'].startswith('Office') and row['theme']['sha256']!=source['theme']['sha256'] and row['theme']['fontScheme']!=source['theme']['fontScheme'], 'theme unchanged'
        else:
            assert row['theme']['name']==source['theme']['name'] and row['theme']['fontScheme']==source['theme']['fontScheme'], 'font scheme changed'
            assert row['theme']['sha256']=='45fd4ae0d4ed96634406c7e5e39afa1102f5a94e5dbdc44cd52889cab01efa49', 'unexpected normalized theme bytes'
    return {'sha256':edit.fresh.sha(file),'controls':rows}

def compact_traces(traces):
    result=[]
    for t in traces:
        text=''.join(chr(c[0]) for c in t['chars'])
        if result and result[-1][0]==t['font']:result[-1][1]+=text
        else:result.append([t['font'],text])
    return result

def inspect_pdf(file,expected,stage,baseline=None):
    doc=fitz.open(file);assert len(doc)==1;page=doc[0]
    assert abs(page.rect.width-720)<.1 and abs(page.rect.height-260)<.1
    label=''.join(s['text'] for b in page.get_text('dict')['blocks'] if 'lines' in b for line in b['lines'] for s in line['spans'] if s['bbox'][1]<60).strip()
    assert label=='combined-scripts'
    fonts={f['id']:f['postscriptName'] for f in expected['fonts']}
    latin='Arial-BoldMT' if stage=='bold' else 'Arial-ItalicMT' if stage=='italic' else 'ArialMT'
    wanted=[[latin,'Latin ' if stage=='source' else 'LatinX ']]
    if stage in ['bold','italic']:wanted=[[latin,'LatinX'],['ArialMT',' ']]
    wanted += [[fonts[k],text] for k,text in [('sc','骨'),('tc','骨'),('ja','骨'),('ko','한'),('ar','لا '),('he','שָׁ '),('greek','Ω')]]
    traces=[t for t in page.get_texttrace() if t['bbox'][1]>60];glyphs=[c for t in traces for c in t['chars']]
    assert compact_traces(traces)==wanted, 'actual PDF face/text sequence'
    assert glyphs and all(c[1]!=0 for c in glyphs)
    # Record actual screen order independently from source/content-stream order.
    positions={}
    for trace in traces:
        for char in trace['chars']:
            if char[0]!=32:positions[trace['font']]=min(positions.get(trace['font'],float('inf')),char[2][0])
    visual_order=sorted(positions,key=positions.get)
    assert visual_order==[latin,*[fonts[k] for k in ['sc','tc','ja','ko','he','ar','greek']]], 'named native visual order changed'
    resources=mixed.font_resources(doc,page)
    stream=b''.join(doc.xref_stream(x) for x in page.get_contents()).decode('latin1')
    pattern=r'BT\s+(?:[-\d.]+\s+Tc\s+)?([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+Tm\s+/([A-Za-z0-9]+)\s+1\s+Tf'
    matrices=[]
    for match in re.findall(pattern,stream):
        matrix=[float(v) for v in match[:6]]
        if matrix[0]!=200:continue
        assert matrix[1]==matrix[2]==0 and matrix[3]==200, 'unexpected synthetic shear'
        matrices.append({'face':resources[match[6]],'matrix':matrix})
    assert set(m['face'] for m in matrices)==set(t['font'] for t in traces), 'body matrix coverage'
    pix=page.get_pixmap(alpha=False);image=Image.frombytes('RGB',(pix.width,pix.height),pix.samples)
    ink=sum(min(p)<200 for p in image.crop((0,60,720,260)).getdata());assert ink>100
    visible=[]
    for trace in traces:
        text=''.join(chr(c[0]) for c in trace['chars'])
        if not text.strip():continue
        box=trace['bbox'];rect=(max(0,math.floor(box[0])),max(0,math.floor(box[1])),min(720,math.ceil(box[2])),min(260,math.ceil(box[3])))
        painted=sum(min(p)<200 for p in image.crop(rect).getdata())
        assert painted>5, 'nonblank font span invisible'
        visible.append({'font':trace['font'],'text':text,'pixels':painted,'bounds':rect})
    if baseline:
        with fitz.open(baseline) as previous:
            old=previous[0].get_pixmap(alpha=False)
            assert (pix.width,pix.height,pix.samples)==(old.width,old.height,old.samples), 'round raster changed'
    return {'sha256':edit.fresh.sha(file),'compactTraces':wanted,'sourceRunOrder':[r['fontId'] for r in expected['expectedRuns']],'nativeVisualFaceOrder':visual_order,'glyphCount':len(glyphs),'glyphZeroCount':0,'darkPixels':ink,'visibleTraceInk':visible,'bodyMatrices':matrices,'glyphsByFont':[{'font':t['font'],'glyphs':[{'unicode':c[0],'glyph':c[1],'origin':c[2]} for c in t['chars']]} for t in traces],'rasterSha256':hashlib.sha256(pix.samples).hexdigest(),'baselineRasterMatch':bool(baseline)}

def neighbors(row):return [(t['font'],[(c['unicode'],c['glyph']) for c in t['glyphs']]) for t in row['glyphsByFont'] if t['font'] not in ['ArialMT','Arial-BoldMT','Arial-ItalicMT']]

def assert_round_fields(package,initial,stage):
    wanted=copy.deepcopy(initial['controls'])
    if stage=='edit':
        first=wanted[0]['runs'][0];assert first['text']=='LatinX ' and 'dirty' not in first['rawAttributes'] and 'err' not in first['rawAttributes']
        space=copy.deepcopy(first);space['text']=' ';space['rawAttributes']['dirty']='0'
        first['text']='LatinX';first['rawAttributes'].update({'dirty':'0','err':'1'})
        wanted[0]['runs'].insert(1,space)
    assert package['controls']==wanted, 'unexpected native round raw fields'

def inspect(directory):
    expected=json.loads((directory/'expected.json').read_text())
    for font in expected['fonts']:assert edit.fresh.sha(font['file'])==font['sha256']
    source=inspect_package(directory/'combined-fonts.pptx',expected)
    result={'source':{'package':source,'pdf':inspect_pdf(directory/'combined-fonts.pdf',expected,'source')}}
    sheet=Image.new('RGB',(720,260*5),'white')
    for i,stage in enumerate(['source',*STAGES]):
        stem='combined-fonts' if stage=='source' else 'combined-'+stage;pdf=directory/(stem+'.pdf')
        if stage!='source':
            result[stage]={'initial':{'package':inspect_package(directory/(stem+'.pptx'),expected,source['controls'][0],stage),'pdf':inspect_pdf(pdf,expected,stage)},'rounds':[]}
            assert neighbors(result[stage]['initial']['pdf'])==neighbors(result['source']['pdf']), 'neighbor glyph sequence changed'
            for round in range(1,4):
                name=stem+'-round'+str(round);package=inspect_package(directory/(name+'.pptx'),expected,source['controls'][0],stage,round)
                assert_round_fields(package,result[stage]['initial']['package'],stage)
                result[stage]['rounds'].append({'name':name,'package':package,'pdf':inspect_pdf(directory/(name+'.pdf'),expected,stage,baseline=pdf)})
        with fitz.open(pdf) as doc:
            pix=doc[0].get_pixmap(alpha=False);sheet.paste(Image.frombytes('RGB',(pix.width,pix.height),pix.samples),(0,260*i))
    regular=result['edit']['initial']['pdf']['rasterSha256']
    assert all(result[s]['initial']['pdf']['rasterSha256']!=regular for s in ['bold','italic','theme']), 'operation not visible'
    sheet.save(directory/'combined-contact-sheet.png')
    return {'schemaVersion':1,'application':'Microsoft PowerPoint 16.104 (16.104.25121423)','artifactsDirectory':str(directory),'sourceSha256':expected['sourceSha256'],'warnings':expected['warnings'],'results':result,'scope':'one original paragraph with Latin/SC/TC/ja/ko/ar/he and Greek; local Latin insertion/style; changed Office theme; three native rounds each','orderingCaveat':'Source run order, content trace order and observed visual face order are distinct. Named native behavior and preservation are checked; no general Unicode bidi or mark-outline conformance claim.'}

if __name__=='__main__':
    directory=Path(sys.argv[1]);result=inspect(directory)
    (directory/'combined-inspection.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps({'authorRuns':8,'stages':STAGES,'roundsPerStage':3,'nativeArtifacts':16}))
