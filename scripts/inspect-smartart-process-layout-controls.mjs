// Optional controlled attribution of the Process frame-overflow observation.
// Transformed source components remain local research inputs, not public output.
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { readFile, writeFile, mkdtemp } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { inspectPptx } from "../packages/pptx-parser/dist/index.js";
import { observeDiagramGeometry } from "./research/diagram-geometry-observation.mjs";

const prepared = JSON.parse(await readFile(new URL("../docs/testing/smartart-native-presentation-candidate-evidence.json", import.meta.url), "utf8"));
const manifest = path.resolve(process.argv[2] ?? path.join(prepared.directory, "manifest.json"));
const controls = path.resolve(process.argv[3] ?? ".tmp/native-diagram-generated/source-controls/manifest.json");
const executable = process.argv[4] ?? "soffice", run = promisify(execFile);
const directory = await mkdtemp(path.resolve(".tmp/native-diagram-generated/process-layout-controls-"));
const application = (await run(executable, ["--version"], { timeout: 15000 })).stdout.trim();
const program = String.raw`
from pathlib import Path
import zipfile,xml.etree.ElementTree as E,json,hashlib,subprocess,sys,importlib.util
manifest,controls,directory,exe,extractor=sys.argv[1:];out=Path(directory)
candidate=next(r for r in json.loads(Path(manifest).read_text())['records'] if r['family']=='process' and r['case']=='standard' and r['cache']=='absent' and r['presentation']=='layout')
source=next(r for r in json.loads(Path(controls).read_text())['records'] if r['family']=='process')
def sha(file):return hashlib.sha256(file.read_bytes()).hexdigest()
def parts(file):
 with zipfile.ZipFile(file) as z:
  assert z.testzip() is None
  assert sum(n.file_size for n in z.infolist())<=16000000
  return {n:z.read(n) for n in z.namelist()}
sourcefile=Path(controls).parent/source['file'];candidatefile=Path(manifest).parent/candidate['file']
assert sha(sourcefile)==source['sha256'];assert sha(candidatefile)==candidate['sha256']
src=parts(sourcefile);gen=parts(candidatefile);D='{http://schemas.openxmlformats.org/drawingml/2006/diagram}'
source_data=E.fromstring(src['ppt/diagrams/data1.xml'])
variables={p.find(D+'prSet').get('presName'):p.find(D+'prSet').find(D+'presLayoutVars') for p in source_data.find(D+'ptLst') if p.get('type')=='pres' and p.find(D+'prSet').find(D+'presLayoutVars') is not None}
assert set(variables)=={'linearFlow','parTx','desTx'}
variants=[('original-source',src,src['ppt/diagrams/data1.xml'],[]),('fresh-shell-source-data',gen,src['ppt/diagrams/data1.xml'],[]),('generated-baseline',gen,gen['ppt/diagrams/data1.xml'],[])]
for name,selected in [('root-variables',{'linearFlow'}),('title-variables',{'parTx'}),('all-variables',set(variables))]:
 data=E.fromstring(gen['ppt/diagrams/data1.xml'])
 for point in data.find(D+'ptLst'):
  prop=point.find(D+'prSet');key=prop.get('presName')
  if key in selected:prop.append(E.fromstring(E.tostring(variables[key])))
 variants.append((name,gen,E.tostring(data,encoding='utf-8',xml_declaration=True),sorted(selected)))
spec=importlib.util.spec_from_file_location('geometry',extractor);module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
records=[]
for name,shell,data,selected in variants:
 file=out/(name+'.pptx')
 with zipfile.ZipFile(file,'w',compression=zipfile.ZIP_DEFLATED) as z:
  for n,b in {**shell,'ppt/diagrams/data1.xml':data}.items():z.writestr(n,b)
 result=subprocess.run([exe,'-env:UserInstallation='+(out/'profile').as_uri(),'--headless','--convert-to','fodp','--outdir',str(out),str(file)],capture_output=True,text=True,timeout=45)
 assert result.returncode==0,result.stderr
 native=file.with_suffix('.fodp');raw=module.extract(file,native)
 records.append(dict(variant=name,file=file.name,inputSha256=sha(file),nativeSha256=sha(native),copiedPresentationVariables=selected,extracted=raw,log=dict(stdout=result.stdout,stderr=result.stderr)))
assert sha(sourcefile)==source['sha256'];assert sha(candidatefile)==candidate['sha256']
print(json.dumps(dict(sourceSha256=source['sha256'],candidateSha256=candidate['sha256'],records=records)))
`;
const extracted = JSON.parse((await run("python3", ["-c", program, manifest, controls, directory, executable,
  fileURLToPath(new URL("./research/extract-native-diagram-geometry.py", import.meta.url))], { timeout: 300000, maxBuffer: 8 * 1024 * 1024 })).stdout);
for (const record of extracted.records) {
  const inspection = await inspectPptx(await readFile(path.join(directory, record.file)));
  const observation = observeDiagramGeometry(record.extracted.snapshot, inspection.diagramInventory.dataModels[0]);
  assert.equal(record.extracted.groupAssociation.status, "located");
  assert.equal(observation.shapes.length, record.extracted.oracle.length);
  for (const [i, oracle] of record.extracted.oracle.entries()) {
    assert.equal(observation.shapes[i].geometryStatus, oracle.status);
    if (oracle.boundsPoints) for (const key of ["x", "y", "width", "height"]) assert.ok(Math.abs(observation.shapes[i].boundsPoints[key] - oracle.boundsPoints[key]) < 1e-8);
  }
  Object.assign(record, { observation, independentGeometryChecked: true }); delete record.extracted;
}
const report = { profile: "native-process-layout-controls-v1", application, directory, manifest, controls, ...extracted,
  limitations: ["Six controlled FODP observations in one named development runtime, not a PowerPoint result or universal attribution.", "The original source and fresh/source-data controls distinguish candidate-only problems from observed producer behavior; equal failure counts do not prove identical internals or geometry.", "Copied presentation variables are diagnostic interventions only; no speculative production writer change is made.", "Third-party source/components and transformed outputs remain local pending licensing admission."] };
await writeFile(new URL("../docs/testing/smartart-process-layout-controls.json", import.meta.url), JSON.stringify(report, null, 2) + "\n");
console.log(JSON.stringify({ directory, cases: report.records.map(r => ({ variant: r.variant, outsideFrame: r.observation.summary.outsideFrame })) }, null, 2));
