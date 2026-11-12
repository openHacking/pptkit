// Optional independent XML and named headless native observations; no GUI edits.
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { platform, release, arch } from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { createPresentation, editTable } from "../packages/core/dist/index.js";
import { resolveLayout } from "../packages/layout/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { tableStyleDeck } from "../tests/fixtures/table-style.mjs";
if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-table-style.mjs [soffice-executable]");
const execute = promisify(execFile), executable = process.argv[2] ?? "soffice", options = { timeout: 45000, maxBuffer: 8 * 1024 * 1024 };
const application = (await execute(executable, ["--version"], options)).stdout.trim(), hash = bytes => createHash("sha256").update(bytes).digest("hex");
await mkdir(".tmp/table-style-native", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/table-style-native/run-")), nativeDirectory = path.join(directory,"native"), pdfDirectory = path.join(directory,"pdf");
await mkdir(nativeDirectory); await mkdir(pdfDirectory);
const expected = [], files = [], logs = [], hashes = [];
for (const name of ["rgb","theme","edited"]) {
  const document = tableStyleDeck(createPresentation,{theme:name === "theme"});
  if (name === "edited") { const slide = document.slides[0], original = slide.elements[0]; slide.replaceElement(original.id,editTable(original,[{type:"insert-row",index:3,height:30}],{frame:"fit-grid"}).table); }
  const layout = resolveLayout(document), generated = await generatePptx(document), file = path.join(directory,`${name}.pptx`); await writeFile(file,generated.bytes); files.push(file); hashes.push({name,sourceSha256:hash(generated.bytes)}); expected.push({name,theme:layout.theme,table:layout.slides[0].elements[0],warnings:generated.warnings});
}
const expectedFile = path.join(directory,"expected.json"); await writeFile(expectedFile,JSON.stringify(expected));
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory,"profile")).href}`;
logs.push((await execute(executable,[profile,"--headless","--convert-to","pptx:Impress Office Open XML","--outdir",nativeDirectory,...files],options)).stdout);
const nativeFiles = expected.map(e=>path.join(nativeDirectory,`${e.name}.pptx`));
logs.push((await execute(executable,[profile,"--headless","--convert-to","pdf:impress_pdf_Export","--outdir",pdfDirectory,...nativeFiles],options)).stdout);
const python = String.raw`
import sys,json,zipfile,xml.etree.ElementTree as E,copy,fitz,os
expectedFile,directory=sys.argv[1:];expected=json.load(open(expectedFile));A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
aliases={'bg1':'lt1','tx1':'dk1','bg2':'lt2','tx2':'dk2'};roles={'background1':'lt1','text1':'dk1','background2':'lt2','text2':'dk2','hyperlink':'hlink','followedHyperlink':'folHlink'}
def package(file):
  with zipfile.ZipFile(file) as z:return {n:z.read(n) for n in z.namelist() if not n.endswith('/')}
def colors(parts):
  scheme=E.fromstring(parts['ppt/theme/theme1.xml']).find('.//'+A+'clrScheme');out={}
  for n in scheme:
    c=list(n)[0];out[n.tag.split('}')[-1]]=(c.get('val') if c.tag==A+'srgbClr' else c.get('lastClr')).upper()
  return out
def actualColor(node,palette):
  if node is None:return None
  if node.tag==A+'srgbClr':return node.get('val').upper()
  if node.tag==A+'schemeClr':return palette.get(aliases.get(node.get('val'),node.get('val')))
  return None
def wantedColor(value,palette):return value.upper().lstrip('#') if isinstance(value,str) else palette[roles.get(value['theme'],value['theme'])]
def inspect(parts,root,entry,strict):
  t=entry['table'];palette=colors(parts);native=root.find('.//'+A+'tbl');rows=native.findall(A+'tr');cols=native.find(A+'tblGrid').findall(A+'gridCol');diff=[]
  if len(rows)!=len(t['rows']) or len(cols)!=len(t['columns']):
    assert not strict
    return {'status':'partial','reason':'native-grid-rewritten','differences':[]}
  for at in t['cellLayout']:
    cell=t['rows'][at['rowIndex']]['cells'][at['cellIndex']];node=rows[at['rowIndex']].findall(A+'tc')[at['columnIndex']];props=node.find(A+'tcPr');body=node.find(A+'txBody');fill=props.find(A+'solidFill') if props is not None else None
    got=actualColor(list(fill)[0],palette) if fill is not None else None;want=wantedColor(cell['style']['fill']['color'],palette) if cell['style']['fill']['type']=='solid' else None
    if got!=want:diff.append({'at':at,'property':'fill','expected':want,'actual':got})
    runs=body.findall('.//'+A+'r');wanted=[r for p in cell['content'] for r in p['runs']]
    # Empty styled cells may be normalized by the producer to endParaRPr.
    gotText=[''.join(r.find(A+'t').itertext()) if r.find(A+'t') is not None else '' for r in runs];wantText=[r['text'] for r in wanted]
    if ''.join(gotText)!=''.join(wantText):diff.append({'at':at,'property':'text','expected':wantText,'actual':gotText})
    if strict:
      assert len(runs)==len(wanted)
      for n,w in zip(runs,wanted):
        p=n.find(A+'rPr');s=w['style'];assert int(p.get('sz'))==round(s['fontSize']*100);assert (p.get('b')=='1')==s['bold'];assert (p.get('i')=='1')==s['italic'];assert actualColor(list(p.find(A+'solidFill'))[0],palette)==wantedColor(s['color'],palette)
      for key,side in [('marL','left'),('marR','right'),('marT','top'),('marB','bottom')]:assert int(props.get(key))==round(cell['style']['margin'][side]*12700)
  if strict:assert not diff,diff
  return {'status':'pass' if not diff else 'partial','checkedAnchors':len(t['cellLayout']),'differences':diff}
result=[]
for entry in expected:
  name=entry['name'];source=package(os.path.join(directory,name+'.pptx'));native=package(os.path.join(directory,'native',name+'.pptx'));root=E.fromstring(source['ppt/slides/slide1.xml']);src=inspect(source,root,entry,True)
  damaged=copy.deepcopy(root);first=damaged.find('.//'+A+'tbl').findall(A+'tr')[0].findall(A+'tc')[0].find(A+'tcPr').find(A+'solidFill');first.clear();E.SubElement(first,A+'srgbClr',{'val':'FF00FF'})
  try:inspect(source,damaged,entry,True)
  except AssertionError:rejected=True
  else:rejected=False
  assert rejected
  observation=inspect(native,E.fromstring(native['ppt/slides/slide1.xml']),entry,False)
  pdf=fitz.open(os.path.join(directory,'pdf',name+'.pdf'));assert len(pdf)==1;page=pdf[0];text=page.get_text();required=['R0C0','Merged body','Upper','Lower','Local','R4C0'];missing=[s for s in required if s not in text];assert not missing,missing
  png=name+'-native.png';page.get_pixmap(matrix=fitz.Matrix(2,2)).save(os.path.join(directory,png))
  result.append({'name':name,'source':src,'deliberateFillMutationRejected':rejected,'nativeFillAndText':observation,'savedPdfText':{'status':'pass','required':required},'png':png})
print(json.dumps({'verified':True,'oracle':'Python ZIP/ElementTree source effective fill/text/style/margin + native fill/text observations + saved PDF text/PNG','scenes':result}))
`;
const independent = JSON.parse((await execute("python3",["-c",python,expectedFile,directory],options)).stdout);
for (const item of hashes) { item.nativeSha256 = hash(await readFile(path.join(nativeDirectory,`${item.name}.pptx`))); item.pdfSha256 = hash(await readFile(path.join(pdfDirectory,`${item.name}.pdf`))); }
const evidence = {schemaVersion:1,recordedAt:new Date().toISOString(),application,platform:`${platform()} ${release()} ${arch()}`,artifactsDirectory:path.relative(process.cwd(),directory),scope:"Materialized table region styles: independent source effective formatting and native saved fill/text observations. Does not establish reusable native table-style rules, source-style writes, full width/alpha/transform parity or interactive editing.",hashes,independent,logs,interactiveEditing:"untested",nativeReusableRules:"unsupported",sourceStyleWrites:"blocked"};
await writeFile(path.join(directory,"evidence.json"),JSON.stringify(evidence,null,2)+"\n"); await writeFile("docs/testing/table-style-evidence.json",JSON.stringify(evidence,null,2)+"\n");
console.log(JSON.stringify({directory,application,scenes:independent.scenes.map(e=>({name:e.name,native:e.nativeFillAndText.status,pdf:e.savedPdfText.status}))}));
