// Optional independent native/PDF border observations. No interactive editing or source-style write admission.
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { platform, release, arch } from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { createPresentation } from "../packages/core/dist/index.js";
import { resolveLayout } from "../packages/layout/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { tableBorderDeck } from "../tests/fixtures/table-borders.mjs";

if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-table-borders.mjs [soffice-executable]");
const execute = promisify(execFile), executable = process.argv[2] ?? "soffice", options = { timeout: 45000, maxBuffer: 8 * 1024 * 1024 };
const application = (await execute(executable, ["--version"], options)).stdout.trim(), hash = bytes => createHash("sha256").update(bytes).digest("hex");
await mkdir(".tmp/table-borders-native", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/table-borders-native/run-")), nativeDirectory = path.join(directory, "native"), pdfDirectory = path.join(directory, "pdf"), controlDirectory = path.join(directory, "source-pdf");
await mkdir(nativeDirectory); await mkdir(pdfDirectory); await mkdir(controlDirectory);
const document = tableBorderDeck(createPresentation), layout = resolveLayout(document), source = await generatePptx(document), svg = await renderPresentationToSvg(document);
const sourceFile = path.join(directory, "source.pptx"), expectedFile = path.join(directory, "expected.json");
await writeFile(sourceFile, source.bytes); await writeFile(expectedFile, JSON.stringify({ theme: layout.theme, tables: layout.slides.map(s => s.elements[0]) }));
for (let i = 0; i < svg.slides.length; i++) await writeFile(path.join(directory, `source-${i + 1}.svg`), svg.slides[i].svg);
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, "profile")).href}`;
const nativeLog = await execute(executable, [profile, "--headless", "--convert-to", "pptx:Impress Office Open XML", "--outdir", nativeDirectory, sourceFile], options);
const nativeFile = path.join(nativeDirectory, "source.pptx"), pdfLog = await execute(executable, [profile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", pdfDirectory, nativeFile], options), pdfFile = path.join(pdfDirectory, "source.pdf");
const controlLog = await execute(executable, [profile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", controlDirectory, sourceFile], options), controlPdf = path.join(controlDirectory, "source.pdf");
const python = String.raw`
import sys,json,zipfile,xml.etree.ElementTree as E,copy,math,fitz,re,os
source,native,pdf,controlPdf,expectedFile,directory=sys.argv[1:];expected=json.load(open(expectedFile));A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
aliases={'bg1':'lt1','tx1':'dk1','bg2':'lt2','tx2':'dk2'}
roles={'background1':'lt1','text1':'dk1','background2':'lt2','text2':'dk2','hyperlink':'hlink','followedHyperlink':'folHlink'}
def package(file):
  with zipfile.ZipFile(file) as z:return {n:z.read(n) for n in z.namelist() if not n.endswith('/')}
def palette(parts):
  scheme=E.fromstring(parts['ppt/theme/theme1.xml']).find('.//'+A+'clrScheme');out={}
  for n in scheme:
    child=list(n)[0];out[n.tag.split('}')[-1]]=child.get('val') if child.tag==A+'srgbClr' else child.get('lastClr')
  return out
def color(value,palette):return value.upper().lstrip('#') if isinstance(value,str) else palette[roles.get(value['theme'],value['theme'])].upper()
def expectedStroke(style,table,colors):
  p=style['paint'];return {'width':style['width'],'dash':style['dash'],'color':color(p['color'],colors) if p['type']=='solid' else None,'alpha':p.get('opacity',1)*table['opacity'] if p['type']=='solid' else 0}
def line(node,colors):
  if node is None:return None
  fill=node.find(A+'solidFill');rgb=None;alpha=0
  if fill is not None:
    c=list(fill)[0];rgb=c.get('val') if c.tag==A+'srgbClr' else colors.get(aliases.get(c.get('val'),c.get('val')))
    a=c.find(A+'alpha');alpha=int(a.get('val'))/100000 if a is not None else 1
  dash=node.find(A+'prstDash');d=dash.get('val') if dash is not None else 'solid'
  return {'width':int(node.get('w','12700'))/12700,'dash':'dot' if d=='sysDot' else d,'color':rgb.upper() if rgb else None,'alpha':alpha}
def equal(a,b,tolerance=0.00004):
  return a is not None and abs(a['width']-b['width'])<=tolerance and a['dash']==b['dash'] and a['color']==b['color'] and abs(a['alpha']-b['alpha'])<=0.00002
def inspect(root,table,colors,strict):
  g=table['borderGrid'];native=root.find('.//'+A+'tbl')
  if native is None:
    assert not strict
    return {'checkedPhysicalEdges':0,'status':'partial','reason':'native-table-unavailable','differences':[]}
  rows=native.findall(A+'tr');cols=native.find(A+'tblGrid').findall(A+'gridCol')
  if len(rows)!=g['rowCount'] or len(cols)!=g['columnCount']:
    assert not strict
    return {'checkedPhysicalEdges':0,'status':'partial','reason':'native-grid-rewritten','expectedGrid':[g['rowCount'],g['columnCount']],'nativeGrid':[len(rows),len(cols)],'differences':[]}
  errors=[];physical=[];owner=[[None]*len(cols) for _ in rows];checked=0
  for r,row in enumerate(rows):
    cells=row.findall(A+'tc');assert len(cells)==len(cols);physical.append([])
    for c,cell in enumerate(cells):
      at=owner[r][c]
      if at is None:
        rs=int(cell.get('rowSpan','1'));cs=int(cell.get('gridSpan','1'));assert rs>0 and cs>0 and r+rs<=len(rows) and c+cs<=len(cols);at=(r,c,rs,cs)
        for y in range(r,r+rs):
          for x in range(c,c+cs):assert owner[y][x] is None;owner[y][x]=at
      assert (cell.get('hMerge') in ('1','true'))==(c>at[1]) and (cell.get('vMerge') in ('1','true'))==(r>at[0])
      if strict:
        assert int(cell.get('rowSpan','1'))==(at[2] if r==at[0] else 1)
        assert int(cell.get('gridSpan','1'))==(at[3] if c==at[1] else 1)
      props=cell.find(A+'tcPr');actual={}
      edges=[('lnL',g['vertical'][r*(len(cols)+1)+c]),('lnR',g['vertical'][r*(len(cols)+1)+c+1]),('lnT',g['horizontal'][r*len(cols)+c]),('lnB',g['horizontal'][(r+1)*len(cols)+c])]
      if (r,c)==at[:2] and (at[2]>1 or at[3]>1):
        sides=[[g['vertical'][(r+i)*(len(cols)+1)+c] for i in range(at[2])],[g['vertical'][(r+i)*(len(cols)+1)+c+at[3]] for i in range(at[2])],[g['horizontal'][r*len(cols)+c+i] for i in range(at[3])],[g['horizontal'][(r+at[2])*len(cols)+c+i] for i in range(at[3])]]
        edges=[(tag,side[0] if all(s==side[0] for s in side) else dict(side[0],width=0,paint={'type':'none'})) for (tag,_),side in zip(edges,sides)]
      for tag,style in edges:
        actual[tag]=line(props.find(A+tag) if props is not None else None,colors);want=expectedStroke(style,table,sp);checked+=1
        if not equal(actual[tag],want,0.1 if not strict else 0.00004):errors.append({'row':r,'column':c,'side':tag,'expected':want,'actual':actual[tag]})
      physical[-1].append(actual)
  if strict:
    assert not errors,errors
    for r in range(len(rows)):
      for c in range(len(cols)):
        if r+1<len(rows) and owner[r][c]!=owner[r+1][c] and owner[r][c][2:]==(1,1) and owner[r+1][c][2:]==(1,1):assert physical[r][c]['lnB']==physical[r+1][c]['lnT']
        if c+1<len(cols) and owner[r][c]!=owner[r][c+1] and owner[r][c][2:]==(1,1) and owner[r][c+1][2:]==(1,1):assert physical[r][c]['lnR']==physical[r][c+1]['lnL']
  return {'checkedPhysicalEdges':checked,'status':'pass' if not errors else 'partial','differences':errors}
s,n=package(source),package(native);sp,np=palette(s),palette(n);observations=[]
for role,value in expected['theme']['colors'].items():assert sp[roles.get(role,role)].upper()==value.upper().lstrip('#')
for i,t in enumerate(expected['tables']):
  part='ppt/slides/slide%d.xml'%(i+1);root=E.fromstring(s[part]);src=inspect(root,t,sp,True)
  damaged=copy.deepcopy(root);first=damaged.find('.//'+A+'lnL');first.set('w',str(int(first.get('w','0'))+12700))
  try:inspect(damaged,t,sp,True)
  except AssertionError:mutationRejected=True
  else:mutationRejected=False
  assert mutationRejected
  observations.append({'source':src,'unapprovedWidthMutationRejected':mutationRejected,'nativeDirectProperties':inspect(E.fromstring(n[part]),t,np,False)})
def transformed(x,y,t):
  b=t['box'];cx=b['x']+b['width']/2;cy=b['y']+b['height']/2;tr=t['transform'];dx=(x-cx)*(-1 if tr['flipH'] else 1);dy=(y-cy)*(-1 if tr['flipV'] else 1);r=math.radians(tr['rotation'])
  return (cx+dx*math.cos(r)-dy*math.sin(r),cy+dx*math.sin(r)+dy*math.cos(r))
def expectedLines(t):
  xs=[t['box']['x']];ys=[t['box']['y']]
  for w in t['columns']:xs.append(xs[-1]+w)
  for row in t['rows']:ys.append(ys[-1]+row['height'])
  g=t['borderGrid'];out=[]
  for axis,styles in [('h',g['horizontal']),('v',g['vertical'])]:
    for i,style in enumerate(styles):
      p=style['paint']
      if p['type']!='solid' or p['opacity']==0 or style['width']==0:continue
      if axis=='h':r,c=divmod(i,g['columnCount']);a=(xs[c],ys[r]);b=(xs[c+1],ys[r])
      else:r,c=divmod(i,g['columnCount']+1);a=(xs[c],ys[r]);b=(xs[c],ys[r+1])
      st=expectedStroke(style,t,sp);out.append({'start':transformed(*a,t),'end':transformed(*b,t),**st})
  return out
def covered(segment,candidates,tolerance=0.25,checkWidth=True,checkAlpha=True):
  ax,ay=segment['start'];bx,by=segment['end'];dx=bx-ax;dy=by-ay;length=math.hypot(dx,dy);intervals=[]
  if length==0:return True
  for other in candidates:
    if other['color']!=segment['color'] or (checkWidth and abs(other['width']-segment['width'])>0.1) or (checkAlpha and abs(other['alpha']-segment['alpha'])>0.01):continue
    points=[other['start'],other['end']]
    if any(abs((x-ax)*dy-(y-ay)*dx)/length>tolerance for x,y in points):continue
    projection=sorted(((x-ax)*dx+(y-ay)*dy)/(length*length) for x,y in points);intervals.append((max(0,projection[0]),min(1,projection[1])))
  reached=0
  for lo,hi in sorted(intervals):
    if hi<lo or lo>reached+tolerance/length:continue
    reached=max(reached,hi)
  return reached>=1-tolerance/length
def observePdf(file,prefix):
  d=fitz.open(file);assert len(d)==len(expected['tables']);out=[]
  for i,page in enumerate(d):
    actual=[]
    for drawing in page.get_drawings():
      if drawing.get('color') is None or 's' not in drawing['type']:continue
      rgb=''.join('%02X'%round(v*255) for v in drawing['color']);st={'width':drawing['width'],'color':rgb,'alpha':drawing.get('stroke_opacity',1),'dashPattern':drawing['dashes']}
      for item in drawing['items']:
        if item[0]=='l':actual.append({'start':tuple(item[1]),'end':tuple(item[2]),**st})
        elif item[0]=='re':
          r=item[1];points=[(r.x0,r.y0),(r.x1,r.y0),(r.x1,r.y1),(r.x0,r.y1),(r.x0,r.y0)]
          for a,b in zip(points,points[1:]):actual.append({'start':a,'end':b,**st})
    wanted=expectedLines(expected['tables'][i]);missing=[j for j,edge in enumerate(wanted) if not covered(edge,actual)];extra=[j for j,edge in enumerate(actual) if not covered(edge,wanted)]
    geometric=[j for j,edge in enumerate(wanted) if covered(edge,actual,checkWidth=False,checkAlpha=False)]
    image=os.path.join(directory,prefix+'-page-%d.png'%(i+1));page.get_pixmap(matrix=fitz.Matrix(2,2)).save(image)
    out.append({'page':i+1,'status':'pass' if not missing and not extra else 'partial','expectedVisibleUnitEdges':len(wanted),'observedStraightSegments':len(actual),'unmatchedExpectedIndices':missing,'unmatchedObservedIndices':extra,'geometryColorCoveredIndices':geometric,'expectedLines':wanted,'observedSegments':actual,'text':page.get_text(),'nativePng':os.path.basename(image),'dashPatterns':sorted(set(e['dashPattern'] for e in actual))})
  assert 'DirectA' in re.sub(r'\s+','',out[0]['text']);assert 'Widemergedcontrol' in re.sub(r'\s+','',out[2]['text'])
  return out
pdfRows,controlRows=observePdf(pdf,'native'),observePdf(controlPdf,'source-control')
print(json.dumps({'verified':True,'oracle':'Python zipfile + ElementTree physical edges/Office merge spans + PyMuPDF straight-line coverage and PNG','sourceTables':observations,'pdf':{'tolerancePoints':0.25,'widthTolerancePoints':0.1,'opacityTolerance':0.01,'dashPixelPatternVerified':False,'pages':pdfRows,'sourceControlPages':controlRows}},ensure_ascii=False))
`;
const independent = JSON.parse((await execute("python3", ["-c", python, sourceFile, nativeFile, pdfFile, controlPdf, expectedFile, directory], options)).stdout);
const evidence = { schemaVersion: 1, recordedAt: new Date().toISOString(), application, platform: `${platform()} ${release()} ${arch()}`, artifactsDirectory: path.relative(process.cwd(), directory), scope: "Direct per-edge table authoring and deterministic shared physical border grids. Independent source XML, native direct-property observations and saved-output PDF line/PNG observations; not full style inheritance, interactive editability, schema certification or source style preservation writes.",
  sourceSha256: hash(source.bytes), nativeSha256: hash(await readFile(nativeFile)), pdfSha256: hash(await readFile(pdfFile)), sourceControlPdfSha256: hash(await readFile(controlPdf)), independent, nativeLog, pdfLog, controlLog, interactiveEditing: "not-performed", sourceStyleSave: "not-performed" };
await writeFile("docs/testing/table-borders-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
await writeFile(path.join(directory, "evidence.json"), JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ directory, application, nativeDirectProperties: independent.sourceTables.map(t => t.nativeDirectProperties.status), pdfLines: independent.pdf.pages.map(p => p.status) }));
