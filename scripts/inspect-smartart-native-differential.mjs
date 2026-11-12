// Optional local, controlled Cycle differential research. These transformed
// third-party inputs are diagnostic artifacts, not preservation-save output.
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { readFile, writeFile, mkdtemp } from "node:fs/promises";
import path from "node:path";

const evidence = JSON.parse(await readFile(new URL("../docs/testing/smartart-native-presentation-candidate-evidence.json", import.meta.url), "utf8"));
const manifest = path.resolve(process.argv[2] ?? path.join(evidence.directory, "manifest.json"));
const controls = path.resolve(process.argv[3] ?? ".tmp/native-diagram-generated/source-controls/manifest.json");
const executable = process.argv[4] ?? "soffice";
const directory = await mkdtemp(path.resolve(".tmp/native-diagram-generated/differential-"));
const run = promisify(execFile);
const application = (await run(executable, ["--version"], { timeout: 15000 })).stdout.trim();
const program = String.raw`
import sys,json,zipfile,re,hashlib,subprocess,fitz
from pathlib import Path
manifest,controls,directory,exe=sys.argv[1:];out=Path(directory)
candidate=next(r for r in json.loads(Path(manifest).read_text())['records'] if r['family']=='cycle' and r['case']=='standard' and r['cache']=='absent')
source=next(r for r in json.loads(Path(controls).read_text())['records'] if r['family']=='cycle')
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def parts(p):
 with zipfile.ZipFile(p) as z:
  assert z.testzip() is None
  return {n:z.read(n) for n in z.namelist()}
cp=Path(manifest).parent/candidate['file'];sp=Path(controls).parent/source['file']
assert sha(cp)==candidate['sha256'];assert sha(sp)==source['sha256']
gen=parts(cp);src=parts(sp);data=gen['ppt/diagrams/data1.xml'].decode()
# A single attribute intervention, keeping all point/edge identities and text.
unlocked=data.replace(' custT="1"','');locked=re.sub(r'(<dgm:prSet phldr="0")',r'\1 custT="1"',unlocked)
attrs='modelId|srcId|destId|cxnId|parTransId|sibTransId|presAssocID'
guid=re.sub(r'('+attrs+r')="(\d+)"',lambda m:m[1]+'="{00000000-0000-4000-8000-'+format(int(m[2]),'012X')+'}"',unlocked)
variants=[('fresh-unlocked',gen,unlocked.encode(),candidate),('fresh-custom-text',gen,locked.encode(),candidate),('fresh-guid',gen,guid.encode(),candidate),('source-shell-generated-data',src,unlocked.encode(),candidate),('fresh-shell-source-data',gen,src['ppt/diagrams/data1.xml'],source)]
results=[]
profile='-env:UserInstallation='+(out/'profile').as_uri()
for name,shell,model,expected in variants:
 file=out/(name+'.pptx')
 with zipfile.ZipFile(file,'w',compression=zipfile.ZIP_DEFLATED) as z:
  for n,b in {**shell,'ppt/diagrams/data1.xml':model}.items():z.writestr(n,b)
 record=dict(variant=name,inputSha256=sha(file),markers=[p for n in expected['authored']['nodes'] for p in n['paragraphs'] if p.strip()],operations=[])
 def command(args):
  p=subprocess.run([exe,profile,'--headless']+args,capture_output=True,text=True,timeout=45)
  assert p.returncode==0,p.stderr
  return dict(stdout=p.stdout,stderr=p.stderr)
 for stage,input in [('direct',file),('reopened',out/'saved'/file.name)]:
  if stage=='reopened':
   saved=out/'saved';saved.mkdir(exist_ok=True)
   record['saveLog']=command(['--convert-to','pptx:Impress Office Open XML','--outdir',str(saved),str(file)])
  dest=out/stage;dest.mkdir(exist_ok=True)
  log=command(['--convert-to','pdf:impress_pdf_Export','--outdir',str(dest),str(input)])
  pdf=dest/file.with_suffix('.pdf').name;document=fitz.open(pdf)
  text='\n'.join(page.get_text() for page in document)
  # Retain a reviewable picture as well as text, without certifying geometry.
  picture=dest/file.with_suffix('.png').name
  document[0].get_pixmap(matrix=fitz.Matrix(1,1)).save(str(picture))
  record['operations'].append(dict(stage=stage,pptxSha256=sha(input),pdfSha256=sha(pdf),picture=str(picture),pictureSha256=sha(picture),pdfExtractedText=text,pages=len(document),log=log))
 results.append(record)
assert sha(cp)==candidate['sha256'];assert sha(sp)==source['sha256']
print(json.dumps(dict(sourceSha256=source['sha256'],candidateSha256=candidate['sha256'],cases=results)))
`;
const output = await run("python3", ["-c", program, manifest, controls, directory, executable], { timeout: 300000, maxBuffer: 8 * 1024 * 1024 });
const { pdfTextEvidence } = await import("./research/pdf-text-evidence.mjs");
const result = JSON.parse(output.stdout);
for (const row of result.cases) for (const operation of row.operations) Object.assign(operation, pdfTextEvidence(operation.pdfExtractedText, row.markers));
const report = { profile: "native-diagram-cycle-differential-v1", application, directory, manifest, controls, ...result,
  limitations: ["One Cycle layout only; producer observations do not establish PowerPoint behavior or native editability.", "PDF lexical matches do not certify distinct occurrences, legibility, layout fidelity or cache freshness.", "Only component/data substitution and the documented attribute/identity interventions are tested; borrowed resources stay local."] };
await writeFile(new URL("../docs/testing/smartart-native-differential-evidence.json", import.meta.url), JSON.stringify(report, null, 2) + "\n");
console.log(JSON.stringify({ directory, cases: report.cases.map(c => ({ variant: c.variant, operations: c.operations.map(o => ({ stage: o.stage, text: o.textMatchStatus })) })) }, null, 2));
