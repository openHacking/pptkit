// Independent source-path/color checks and named headless source-style observations.
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdir,mkdtemp,readFile,writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { platform,release,arch } from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { createPresentation,normalizePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
import { tableStyleFixture,paint,region,TABLE_STYLE_ID } from "../packages/pptx-parser/test/table-style-fixture.mjs";
if(process.argv.length>3)throw new Error("Usage: node scripts/inspect-table-style-import.mjs [soffice-executable]");
const execute=promisify(execFile),executable=process.argv[2]??"soffice",options={timeout:45000,maxBuffer:8*1024*1024},hash=bytes=>createHash("sha256").update(bytes).digest("hex");
const application=(await execute(executable,["--version"],options)).stdout.trim();await mkdir(".tmp/table-style-import-native",{recursive:true});const directory=await mkdtemp(path.resolve(".tmp/table-style-import-native/run-")),nativeDir=path.join(directory,"native"),pdfDir=path.join(directory,"pdf"),sourcePdfDir=path.join(directory,"source-pdf");for(const dir of[nativeDir,pdfDir,sourcePdfDir])await mkdir(dir);
const expected=[],sources=[],hashes=[];
for(const name of["regions","theme","local"]){
  const document=createPresentation({size:{width:400,height:300},theme:{colors:{accent1:"3579BD"},fonts:{body:"Arial",heading:"Arial"}}});document.addSlide().addElement({type:"table",box:{x:30,y:30,width:280,height:120},columns:Array(4).fill(70),rows:Array.from({length:4},(_,r)=>({height:30,cells:Array.from({length:4},(_,c)=>({content:[{runs:[{text:`R${r}C${c}`,style:{fontFamily:"Arial",fontSize:10}}]}]}))}))});
  const parts=readGeneratedZip((await generatePptx(document)).bytes),fixture=tableStyleFixture({flags:name==="theme"?"":undefined,...(name==="theme"?{rules:region("wholeTbl",paint("accent1",true))}:{})});
  parts.set("ppt/tableStyles.xml",Buffer.from(fixture.parts["deck/tableStyles.xml"]));
  let slide=parts.get("ppt/slides/slide1.xml").toString().replace(/<a:tblPr[^>]*>[\s\S]*?<\/a:tblPr>/,`<a:tblPr ${name==="theme"?"":'firstRow="1" lastRow="1" firstCol="1" lastCol="1" bandRow="1" bandCol="1"'}><a:tableStyleId>${TABLE_STYLE_ID}</a:tableStyleId></a:tblPr>`).replace(/<a:tcPr>[\s\S]*?<\/a:tcPr>/g,"<a:tcPr/>").replace(/<a:rPr([^>]*)>([\s\S]*?)<\/a:rPr>/g,(_,attrs,children)=>`<a:rPr${attrs.replace(/ [bi]="[^"]*"/g,"")}>${children.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/g,"")}</a:rPr>`);
  if(name==="local")slide=slide.replace("<a:tcPr/>",`<a:tcPr>${paint("FFCC66")}</a:tcPr>`);
  parts.set("ppt/slides/slide1.xml",Buffer.from(slide));const bytes=zip([...parts].map(([name,body])=>({name,body}))).bytes,file=path.join(directory,`${name}.pptx`);await writeFile(file,bytes);sources.push(file);hashes.push({name,sourceSha256:hash(bytes)});
  const session=await importPptx(bytes,{projection:"content",editSaveProfile:"text-runs-v3"});try{expected.push({name,table:normalizePresentation(session.document).slides[0].elements[0],provenance:session.report.tableStyleInheritance});}finally{session.dispose();}
}
const expectedFile=path.join(directory,"expected.json");await writeFile(expectedFile,JSON.stringify(expected));const profile=`-env:UserInstallation=${pathToFileURL(path.join(directory,"profile")).href}`,logs=[];
logs.push((await execute(executable,[profile,"--headless","--convert-to","pptx:Impress Office Open XML","--outdir",nativeDir,...sources],options)).stdout);
logs.push((await execute(executable,[profile,"--headless","--convert-to","pdf:impress_pdf_Export","--outdir",pdfDir,...expected.map(e=>path.join(nativeDir,`${e.name}.pptx`))],options)).stdout);
logs.push((await execute(executable,[profile,"--headless","--convert-to","pdf:impress_pdf_Export","--outdir",sourcePdfDir,...sources],options)).stdout);
const python=String.raw`
import sys,os,json,zipfile,xml.etree.ElementTree as E,copy,fitz,re
expectedFile,directory=sys.argv[1:];expected=json.load(open(expectedFile));A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
def package(file):
 with zipfile.ZipFile(file) as z:return {n:z.read(n) for n in z.namelist() if not n.endswith('/')}
def palette(parts):
 scheme=E.fromstring(parts['ppt/theme/theme1.xml']).find('.//'+A+'clrScheme');out={}
 for n in scheme:
  c=list(n)[0];out[n.tag.split('}')[-1]]=c.get('val') if c.tag==A+'srgbClr' else c.get('lastClr')
 return out
def location(parts,at,root=None):
 root=E.fromstring(parts[at['part'].lstrip('/')]) if root is None else root;pieces=at['path'].split('/')[1:];node=root
 for i,piece in enumerate(pieces):
  match=re.fullmatch(r'([^:]+):([^\[]+)\[(\d+)\]',piece);assert match,piece
  prefix,local,index=match.groups();tag=A+local if prefix=='a' else '{http://schemas.openxmlformats.org/presentationml/2006/main}'+local
  if i==0:assert node.tag==tag and int(index)==1
  else:
   matches=node.findall(tag);assert len(matches)>=int(index),"Selected source declaration is missing";node=matches[int(index)-1]
 return node
def rgb(node,colors):
 if node is None:return None
 if node.tag==A+'srgbClr':return node.get('val').upper()
 if node.tag==A+'schemeClr':return colors.get({'bg1':'lt1','tx1':'dk1','bg2':'lt2','tx2':'dk2'}.get(node.get('val'),node.get('val')))
 return None
def verify(parts,entry):
 colors=palette(parts);count=0
 for p in entry['provenance']['properties']:
  if p['property']!='cell.fill' or p['status']!='mapped':continue
  node=location(parts,p['source']);assert rgb(node,colors)==p['value']['color'];count+=1
 return count
observations=[]
for entry in expected:
 name=entry['name'];s=package(os.path.join(directory,name+'.pptx'));n=package(os.path.join(directory,'native',name+'.pptx'));checked=verify(s,entry)
 damaged=copy.deepcopy(s);part='ppt/tableStyles.xml';root=E.fromstring(damaged[part]);node=root.find('.//'+A+'wholeTbl').find('.//'+A+'srgbClr')
 if node is None:node=root.find('.//'+A+'wholeTbl').find('.//'+A+'schemeClr')
 if node.tag==A+'schemeClr':node.tag=A+'srgbClr'
 node.set('val','FF00FF');damaged[part]=E.tostring(root)
 try:verify(damaged,entry)
 except AssertionError:rejected=True
 else:rejected=False
 # Conditional scenes may select no wholeTbl fills; mutate a selected source too.
 if not rejected:
  selected=next(p for p in entry['provenance']['properties'] if p['property']=='cell.fill' and p['source']['part'].endswith('tableStyles.xml'));root=E.fromstring(s[part]);node=location(s,selected['source'],root);node.tag=A+'srgbClr';node.set('val','FF00FF');damaged[part]=E.tostring(root)
  try:verify(damaged,entry)
  except AssertionError:rejected=True
  else:rejected=False
 assert rejected
 root=E.fromstring(n['ppt/slides/slide1.xml']);rows=root.find('.//'+A+'tbl').findall(A+'tr');colors=palette(n);diff=[]
 for r,row in enumerate(entry['table']['rows']):
  cells=rows[r].findall(A+'tc')
  for c,cell in enumerate(row['cells']):
   props=cells[c].find(A+'tcPr');fill=props.find(A+'solidFill') if props is not None else None;got=rgb(list(fill)[0],colors) if fill is not None else None;want=cell['style']['fill'].get('color')
   if got!=want:diff.append({'row':r,'column':c,'expected':want,'nativeDirectFill':got})
 pdf=[]
 for folder,label in [('pdf','native'),('source-pdf','source-control')]:
  book=fitz.open(os.path.join(directory,folder,name+'.pdf'));assert len(book)==1;page=book[0];text=page.get_text();missing=[f'R{r}C{c}' for r in range(4) for c in range(4) if f'R{r}C{c}' not in text];assert not missing,missing;png=name+'-'+label+'.png';page.get_pixmap(matrix=fitz.Matrix(2,2)).save(os.path.join(directory,png));pdf.append({'kind':label,'text':'pass','png':png})
 observations.append({'name':name,'independentSourceFillProperties':checked,'deliberateColorMutationRejected':rejected,'nativeDirectFills':{'status':'pass' if not diff else 'partial','differences':diff},'pdf':pdf})
print(json.dumps({'verified':True,'oracle':'Python ZIP/ElementTree selected source path and color resolution, deliberate mutation, native direct fills and source-control/saved PDF text/PNG','scenes':observations}))
`;
const independent=JSON.parse((await execute("python3",["-c",python,expectedFile,directory],options)).stdout);
for(const h of hashes){h.nativeSha256=hash(await readFile(path.join(nativeDir,`${h.name}.pptx`)));h.pdfSha256=hash(await readFile(path.join(pdfDir,`${h.name}.pdf`)));h.sourceControlPdfSha256=hash(await readFile(path.join(sourcePdfDir,`${h.name}.pdf`)));}
const evidence={schemaVersion:1,recordedAt:new Date().toISOString(),application,platform:`${platform()} ${release()} ${arch()}`,artifactsDirectory:path.relative(process.cwd(),directory),scope:"Source table-style direct-region profile: selected source colors/provenance and named native saved direct-fill observations. Conditional precedence, native style retention, borders, alpha/fonts/shaping, interactive editing and source-style writes are separate gates.",hashes,independent,logs,interactiveEditing:"untested",sourceStyleWrites:"blocked"};await writeFile(path.join(directory,"evidence.json"),JSON.stringify(evidence,null,2)+"\n");await writeFile("docs/testing/table-style-import-evidence.json",JSON.stringify(evidence,null,2)+"\n");console.log(JSON.stringify({directory,application,scenes:independent.scenes.map(s=>({name:s.name,fills:s.nativeDirectFills.status}))}));
