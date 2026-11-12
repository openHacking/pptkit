import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import { spawnSync } from "node:child_process";
import path from "node:path";
import { inspectPptx, importPptx } from "../packages/pptx-parser/dist/index.js";
import { projectDiagramContent } from "./research/diagram-content-model.mjs";

// Optional offline verifier. UI operations are performed/observed separately;
// this runner never launches PowerPoint or turns file checks into UI passes.
const sourceDirectory = path.resolve(process.argv[2] ?? ".tmp/smartart-source-corpus/0ea8f79d15a6");
const nativeDirectory = path.resolve(process.argv[3] ?? ".tmp/smartart-native-model-20261025");
const provenance = JSON.parse(await readFile(path.join(sourceDirectory, "provenance.json"), "utf8"));
const observations = JSON.parse(await readFile(path.join(nativeDirectory, "observations.json"), "utf8"));
const sha = bytes => createHash("sha256").update(bytes).digest("hex");
const records = [];
for (const [name, family] of [["smartart-accent-process.pptx", "process"], ["smartart-org-chart.pptx", "hierarchy"], ["smartart-cycle.pptx", "cycle"]]) {
  const sample = provenance.samples.find(s => s.name === name); assert.ok(sample);
  const source = await readFile(path.join(sourceDirectory, name)); assert.equal(sha(source), sample.sha256);
  const observation = observations.cases.find(c => c.name === name); assert.ok(observation);
  const versions = [{ kind: "source", file: path.join(sourceDirectory, name), bytes: source }];
  if (observation.nativeSaved === true) versions.push({ kind: "native-saved", file: path.join(nativeDirectory, name), bytes: await readFile(path.join(nativeDirectory, name)) });
  const results = [];
  for (const version of versions) {
    const inspected = await inspectPptx(version.bytes), [diagram] = inspected.slides[0].diagrams;
    const layoutIdentifier = diagram.components.find(c => c.role === "layout").layoutIdentifier;
    const model = inspected.diagramInventory.dataModels.find(m => m.part === diagram.components.find(c => c.role === "data").part);
    const projected = projectDiagramContent(model, layoutIdentifier); assert.equal(projected.candidate.family, family);
    const session = await importPptx(version.bytes, { projection: "structure" });
    try {
      const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready");
      assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(version.bytes));
    } finally { session.dispose(); }
    results.push({ kind: version.kind, file: version.file, sha256: sha(version.bytes), exactNoOp: true, projected, model });
  }
  const artifacts = [];
  for (const artifact of observation.artifacts ?? []) {
    assert.equal(path.basename(artifact), artifact);
    const bytes = await readFile(path.join(nativeDirectory, artifact)); artifacts.push({ file: artifact, sha256: sha(bytes), bytes: bytes.length });
  }
  records.push({ name, family, source: sample.source, sourceSha256: sample.sha256, observation, artifacts, results });
}
await writeFile(path.join(nativeDirectory, "content-inspection.json"), JSON.stringify(records, null, 2) + "\n");
const independentScript = String.raw`
import hashlib,json,sys,zipfile,xml.etree.ElementTree as E
from pathlib import Path
D='{http://schemas.openxmlformats.org/drawingml/2006/diagram}';A='{http://schemas.openxmlformats.org/drawingml/2006/main}';S='{http://schemas.microsoft.com/office/drawing/2008/diagram}'
records=json.loads(Path(sys.argv[1]).read_text()); output=[]
for record in records:
 versions=[]; inventories=[]
 for result in record['results']:
  p=Path(result['file']); assert hashlib.sha256(p.read_bytes()).hexdigest()==result['sha256']
  with zipfile.ZipFile(p) as z:
   assert z.testzip() is None
   model=result['model'];root=E.fromstring(z.read(model['part'].lstrip('/')));raw=list(root.iter());idx={n:i+1 for i,n in enumerate(raw)}
   points=[n for n in root.find(D+'ptLst') if n.tag==D+'pt' and n.get('type','node') in ['doc','node','asst']]
   byid={n.get('modelId'):n for n in points};assert len(byid)==len(points)
   doc=[n for n in points if n.get('type','node')=='doc'];assert len(doc)==1
   adjacency={n.get('modelId'):[] for n in points}
   for c in root.find(D+'cxnLst'):
    if c.get('type','parOf')!='parOf':continue
    adjacency[c.get('srcId')].append((int(c.get('srcOrd')),c.get('destId')))
   for children in adjacency.values():children.sort()
   projected=result['projected'];candidate=projected['candidate'];bindings={b['id']:raw[b['elementIndex']-1] for b in projected['bindings']}
   assert len(candidate['nodes'])==len(points)-1 and len(bindings)==len(points)-1
   assert [bindings[i].get('modelId') for i in candidate['roots']]==[to for order,to in adjacency[doc[0].get('modelId')]]
   content={}
   for node in candidate['nodes']:
    source=bindings[node['id']];key=source.get('modelId');paras=[]
    body=source.find(D+'t')
    if body is not None:
     for para in body.findall(A+'p'):
      text=''
      for run in para:
       if run.tag==A+'r':text+=''.join(run.find(A+'t').itertext())
       elif run.tag==A+'br':text+='\n'
       elif run.tag==A+'fld':raise AssertionError('unadmitted stored field')
      paras.append(text)
    assert node['paragraphs']==paras
    assert node['role']==('assistant' if source.get('type')=='asst' else 'content')
    assert [bindings[i].get('modelId') for i in node['children']]==[to for order,to in adjacency[key]]
    content[key]=dict(role=node['role'],paragraphs=paras,children=[to for order,to in adjacency[key]])
   if record['family']=='cycle':assert candidate['closure']=='layout-generated' and all(not n['children'] for n in candidate['nodes'])
   layouts=[E.fromstring(z.read(n)).get('uniqueId') for n in z.namelist() if n.startswith('ppt/diagrams/layout') and n.endswith('.xml')]
   assert layouts==[projected['source']['layoutIdentifier']]
   caches=[E.fromstring(z.read(n)) for n in z.namelist() if n.startswith('ppt/diagrams/drawing') and n.endswith('.xml')]
   shapes=[n for c in caches for n in c.iter(S+'sp')]
   all_model_ids={n.get('modelId') for n in root.find(D+'ptLst')}
   assert all(n.get('modelId') in all_model_ids for n in shapes)
   inventories.append(content)
   versions.append(dict(kind=result['kind'],sha256=result['sha256'],exactNoOp=result['exactNoOp'],layout=layouts[0],nodes=len(content),assistants=sum(n['role']=='assistant' for n in content.values()),roots=[bindings[i].get('modelId') for i in candidate['roots']],paragraphsBySourceId=content,cacheShapes=len(shapes),independentTopologyTextAndRoles=True))
 if len(versions)==2:
  before,after=inventories;assert set(before)<set(after) and len(after)==len(before)+1
  for key,node in before.items():
   assert after[key]['role']==node['role']
   if record['family']=='process' and node['paragraphs']==['a']:assert after[key]['paragraphs']==['a PPTKit Process']
   elif record['family']=='hierarchy' and node['paragraphs']==['Manager','Second para']:assert after[key]['paragraphs']==['Manager PPTKit Hierarchy','Second para']
   else:assert after[key]['paragraphs']==node['paragraphs']
  new=next(iter(set(after)-set(before)));assert after[new]['paragraphs']==['PPTKit Added '+('Process' if record['family']=='process' else 'Hierarchy')]
  assert versions[-1]['cacheShapes']>versions[0]['cacheShapes']
 output.append(dict(name=record['name'],family=record['family'],source=record['source'],observation=record['observation'],artifacts=record['artifacts'],versions=versions))
print(json.dumps(dict(python=sys.version.split()[0],cases=output)))
`;
const checked = spawnSync("python3", ["-c", independentScript, path.join(nativeDirectory, "content-inspection.json")], { encoding: "utf8", timeout: 60000, maxBuffer: 4 * 1024 * 1024 });
assert.equal(checked.status, 0, checked.stderr);
const independent = JSON.parse(checked.stdout);
const evidence = { profile: "smartart-content-model-research-v1", application: observations.application, sourceCommit: provenance.commit,
  retainedDirectory: nativeDirectory, runtime: { node: process.version, python: independent.python }, cases: independent.cases,
  limitations: ["Candidate content model only; no public Core diagram model, layout engine, native authoring API or permitted diagram source writes.", "Native UI fields are manual observations with retained screenshots/AX hashes; this verifier independently checks files and does not replay or certify UI actions.", "Only the stated native operations/families were observed; unavailable operations remain untested. No general compatibility or visual-fidelity claim.", "Source fixtures remain local pending redistribution/licensing admission; candidate ordinal IDs are projection-local, not stable authoring IDs."] };
await writeFile(new URL("../docs/testing/smartart-content-model-evidence.json", import.meta.url), JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify(evidence.cases.map(c => ({ name: c.name, family: c.family, nativeSaved: c.observation.nativeSaved, versions: c.versions.map(v => ({ kind: v.kind, nodes: v.nodes, cacheShapes: v.cacheShapes })) })), null, 2));
