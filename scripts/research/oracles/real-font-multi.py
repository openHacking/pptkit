"""Five-script native editing experiment; source text and PDF ToUnicode stay distinct."""
import copy, hashlib, importlib.util, json, re, sys
from pathlib import Path
import fitz
from PIL import Image
spec = importlib.util.spec_from_file_location('edit',Path(__file__).with_name('real-font-edit.py'))
edit = importlib.util.module_from_spec(spec); spec.loader.exec_module(edit)
CASES = {
 'map-tc': {'runs':[('骨直令返','zh-TW',None),('返','zh-TW','en-US'),('Ω','el-GR','zh-TW')], 'regular':'STSongti-TC-Regular','bold':'STSongti-TC-Bold','trace':'骨直令返返Ω'},
 'map-ja': {'runs':[('骨直令返','ja',None),('返','zh-CN','en-US'),('Ω','el-GR','ja')], 'regular':'HiraMinProN-W3','bold':'HiraMinProN-W6','trace':'骨直令返返Ω'},
 'map-ko': {'runs':[('한글','ko',None),('글','ko-KR','en-US'),('Ω','el-GR','ko')], 'regular':'AppleSDGothicNeo-Regular','bold':'AppleSDGothicNeo-Bold','trace':'한글글Ω'},
 'map-ar': {'runs':[('السَّلَامُم','ar',None),('Ω','el-GR',None)], 'regular':'GeezaPro','bold':'GeezaPro-Bold','trace':'السَّلامُمΩ','fallback':True},
 'map-he': {'runs':[('שָׁלוֹםם','he',None),('Ω','el-GR',None)], 'regular':'ArialHebrew','bold':'ArialHebrew-Bold','trace':'שָׁלוֹםםΩ','fallback':True},
}
STAGES = ['edit','bold','italic','theme']

def inspect_package(file,source,stage):
 rows = edit.read_package(file)
 for row, original in zip(rows,source):
  assert row['id'] == original['id'], 'order/label'
  case = CASES.get(row['id'])
  assert row['geometry'] == original['geometry'] and row['bodyProperties'] == original['bodyProperties'], 'frame changed'
  assert row['paragraphProperties'] == ([None] if case else original['paragraphProperties']), 'paragraph properties'
  if not case:
   assert [edit.semantic_run(r) for r in row['runs']] == [edit.semantic_run(r) for r in original['runs']], 'unrelated run changed'
  else:
   expected = [{'text':text,'language':lang,'altLanguage':alt,'size':'4800','bold':stage=='bold','italic':stage=='italic','slots':original['runs'][0]['slots']} for text,lang,alt in case['runs']]
   assert [edit.semantic_run(r) for r in row['runs']] == expected, 'edited text/lang/topology/slots/style'
  if stage == 'theme':
   assert row['theme']['sha256'] != original['theme']['sha256'] and row['theme']['fontScheme'] != original['theme']['fontScheme'], 'theme unchanged'
   assert row['theme']['name'].startswith('Office')
  else: assert row['theme'] == original['theme'], 'unexpected theme change'
 return {'sha256':edit.fresh.sha(file),'controls':rows}

def inspect_pdf(file,expected,stage,baseline=None):
 doc=fitz.open(file); assert len(doc)==16
 previous=fitz.open(baseline) if baseline else None
 fonts={f['id']:f['postscriptName'] for f in expected['fonts']}; rows=[]
 for i,(page,control) in enumerate(zip(doc,expected['controls'])):
  assert abs(page.rect.width-720)<.1 and abs(page.rect.height-260)<.1
  label=''.join(s['text'] for b in page.get_text('dict')['blocks'] if 'lines' in b for line in b['lines'] for s in line['spans'] if s['bbox'][1]<60).strip()
  assert label==control['id'], 'PDF order'
  traces=[t for t in page.get_texttrace() if t['bbox'][1]>60]; glyphs=[c for t in traces for c in t['chars']]
  assert glyphs and all(c[1]!=0 for c in glyphs), 'empty/missing glyph'
  actual=sorted(set(t['font'] for t in traces)); case=CASES.get(control['id'])
  row={'id':control['id'],'fonts':actual,'glyphCount':len(glyphs),'glyphZeroCount':0}
  if case:
   target=case['bold'] if stage=='bold' else case['regular']
   fallback=('Cambria-Bold' if stage=='bold' else 'Cambria-Italic' if stage=='italic' else 'Cambria') if case.get('fallback') else target
   assert actual==sorted(set([target,fallback])), 'actual body/fallback face differs'
   for trace in traces:
    for char in trace['chars']: assert trace['font']==(fallback if char[0]==937 else target), 'font per glyph differs'
   text=''.join(chr(c[0]) for c in glyphs)
   assert text==case['trace'], 'named PDF ToUnicode changed'
   source_text=''.join(r[0] for r in case['runs'])
   row.update({'sourceText':source_text,'pdfTraceText':text,'pdfTraceTextEqualsSource':text==source_text,'glyphsByFont':[{'font':t['font'],'glyphs':[{'unicode':c[0],'glyph':c[1],'origin':c[2]} for c in t['chars']]} for t in traces]})
   resource_faces={font[4]:font[3].split('+')[-1] for font in page.get_fonts()}
   stream=b''.join(doc.xref_stream(x) for x in page.get_contents()).decode('latin1')
   pattern=r'BT\s+(?:[-\d.]+\s+Tc\s+)?([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+Tm\s+/([A-Za-z0-9]+)\s+1\s+Tf'
   matrices=[]
   for match in re.findall(pattern,stream):
    face=resource_faces[match[6]]
    if face not in actual: continue
    matrix=[float(v) for v in match[:6]]
    assert matrix[0]==matrix[3]==200 and matrix[1]==0
    shear=stage=='italic' and not face.startswith('Cambria')
    assert abs(matrix[2]/matrix[0]-(67.96875/200 if shear else 0))<1e-9, 'body shear differs'
    matrices.append({'face':face,'matrix':matrix,'syntheticShear':shear})
   assert set(m['face'] for m in matrices)==set(actual), 'missing body matrix'
   row['bodyTextMatrices']=matrices
  else: assert actual==[fonts[control['fontId']]], 'unrelated actual face changed'
  pix=page.get_pixmap(alpha=False); image=Image.frombytes('RGB',(pix.width,pix.height),pix.samples)
  ink=sum(min(p)<200 for p in image.crop((0,60,720,260)).getdata()); assert ink>100
  row.update({'darkPixels':ink,'rasterSha256':hashlib.sha256(pix.samples).hexdigest()})
  if previous:
   old=previous[i].get_pixmap(alpha=False)
   assert (pix.width,pix.height,pix.samples)==(old.width,old.height,old.samples), 'round raster changed'
  rows.append(row)
 return {'sha256':edit.fresh.sha(file),'controls':rows,'baselineRasterMatch':bool(previous)}

def inspect(directory):
 expected=json.loads((directory/'expected.json').read_text()); edit.fresh.inspect_package(directory/'fresh-fonts.pptx',expected,True)
 for font in expected['fonts']: assert edit.fresh.sha(font['file'])==font['sha256']
 source=edit.read_package(directory/'fresh-round3.pptx'); results={}; original=fitz.open(directory/'fresh-fonts.pdf')
 sheet=Image.new('RGB',(720*4,260*5),'white')
 for col,stage in enumerate(STAGES):
  name='fresh-'+stage+'-multi'; pdf=directory/(name+'.pdf')
  results[stage]={'initial':{'package':inspect_package(directory/(name+'.pptx'),source,stage),'pdf':inspect_pdf(pdf,expected,stage)},'rounds':[]}
  with fitz.open(pdf) as doc:
   for row,case in enumerate(CASES):
    index=next(i for i,c in enumerate(expected['controls']) if c['id']==case)
    pix=doc[index].get_pixmap(alpha=False);sheet.paste(Image.frombytes('RGB',(pix.width,pix.height),pix.samples),(col*720,row*260))
   if stage!='theme':
    for i,c in enumerate(expected['controls']):
     if c['id'] not in CASES: assert doc[i].get_pixmap(alpha=False).samples==original[i].get_pixmap(alpha=False).samples, 'unrelated raster changed'
  for round in range(1,4):
   stem=name+'-round'+str(round); package=inspect_package(directory/(stem+'.pptx'),source,stage)
   wanted=copy.deepcopy(results[stage]['initial']['package']['controls'])
   if stage=='edit':
    he=next(c for c in wanted if c['id']=='map-he'); assert 'dirty' not in he['runs'][1]['rawAttributes']; he['runs'][1]['rawAttributes']['dirty']='0'
   assert package['controls']==wanted, 'round raw fields changed'
   results[stage]['rounds'].append({'name':stem,'package':package,'pdf':inspect_pdf(directory/(stem+'.pdf'),expected,stage,baseline=pdf)})
 for id in CASES:
  index=next(i for i,c in enumerate(expected['controls']) if c['id']==id)
  regular=results['edit']['initial']['pdf']['controls'][index]['rasterSha256']
  for stage in ['bold','italic','theme']: assert results[stage]['initial']['pdf']['controls'][index]['rasterSha256']!=regular, 'operation not visible'
 sheet.save(directory/'multi-contact-sheet.png')
 return {'schemaVersion':1,'scope':'five named script pages; append native-script character and Greek, full-body bold/italic, Office theme; three native rounds each','application':'Microsoft PowerPoint 16.104 (16.104.25121423)','artifactsDirectory':str(directory),'sourceSha256':expected['sourceSha256'],'cases':CASES,'results':results,'pdfToUnicodeCaveat':'Arabic trace omits a source fatha; Hebrew uses presentation forms. Exact source text is verified in XML, not inferred from PDF extraction. No complete bidi/mark-shaping certification.'}
if __name__=='__main__':
 directory=Path(sys.argv[1]); result=inspect(directory)
 (directory/'multi-inspection.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
 print(json.dumps({'stages':STAGES,'editedPages':5,'roundsPerStage':3,'pagesPerArtifact':16}))
