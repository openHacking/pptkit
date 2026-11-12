// Optional verifier for generated candidates edited through the native UI.
// File checks do not replay or certify the separately recorded UI operations.
import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { readFile, writeFile, mkdir } from "node:fs/promises";
import path from "node:path";
import { inspectPptx, importPptx, compareOpcPackages } from "../packages/pptx-parser/dist/index.js";
import { projectDiagramContent } from "./research/diagram-content-model.mjs";

const directory = path.resolve(process.argv[2] ?? ".tmp/native-diagram-powerpoint-kmjuxfmg");
const observations = JSON.parse(await readFile(path.join(directory, "observations.json"), "utf8"));
const inputs = JSON.parse(await readFile(path.join(directory, "inputs.json"), "utf8"));
const sha = bytes => createHash("sha256").update(bytes).digest("hex"), run = promisify(execFile), records = [];
const layouts = { process: "process3", hierarchy: "orgChart1", cycle: "cycle2" };
for (const input of inputs) {
  assert.ok(Object.hasOwn(layouts, input.family));
  assert.equal(sha(await readFile(input.input)), input.sha256, "The original generated candidate must stay unchanged.");
}
for (const stage of observations.stages) {
  assert.equal(path.basename(stage.file), stage.file); assert.ok(Object.hasOwn(layouts, stage.family));
  const file = path.join(directory, stage.file), bytes = await readFile(file); assert.equal(sha(bytes), stage.sha256);
  const inspection = await inspectPptx(bytes); assert.equal(inspection.diagramInventory.dataModels.length, 1);
  const diagram = inspection.slides[0].diagrams[0], layout = diagram.components.find(c => c.role === "layout").layoutIdentifier;
  assert.equal(layout, `urn:microsoft.com/office/officeart/2005/8/layout/${layouts[stage.family]}`);
  const model = inspection.diagramInventory.dataModels[0], projected = projectDiagramContent(model, layout);
  assert.deepEqual(model.findings, [], "Named native stages must have a complete recognized source graph.");
  for (const point of model.points) for (const ref of [point.connection, point.presentationAssociation]) if (ref) assert.equal(ref.status, "resolved");
  for (const edge of model.connections) for (const ref of [edge.source, edge.destination, edge.parentTransition, edge.siblingTransition]) if (ref) assert.equal(ref.status, "resolved");
  assert.equal(projected.candidate.nodes.length, stage.expected.nodes);
  assert.equal(projected.candidate.roots.length, stage.expected.roots);
  assert.equal(projected.candidate.nodes.filter(n => n.role === "assistant").length, stage.expected.assistants);
  assert.deepEqual(projected.candidate.nodes.flatMap(n => n.paragraphs), stage.expected.paragraphs);
  for (const mapping of inspection.diagramInventory.cacheMappings) assert.ok(mapping.shapes.every(shape => shape.model.status === "resolved"));
  const session = await importPptx(bytes, { projection: "content", diagramPreview: "cache" });
  try {
    const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready");
    assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
    records.push({ ...stage, file, layout, projected, model, exactNoOp: true, preview: session.report.diagramPreviews,
      cacheShapes: inspection.diagramInventory.cacheMappings.reduce((n, m) => n + m.shapes.length, 0) });
  } finally { session.dispose(); }
}
const artifacts = [];
for (const file of observations.artifacts) {
  assert.equal(path.basename(file), file); const bytes = await readFile(path.join(directory, file));
  artifacts.push({ file, sha256: sha(bytes), bytes: bytes.length });
}
const intermediate = path.join(directory, "generated-native-inspection.json");
await writeFile(intermediate, JSON.stringify(records, null, 2) + "\n");
const python = String.raw`
import sys,json,hashlib,zipfile,xml.etree.ElementTree as E
from pathlib import Path
D='{http://schemas.openxmlformats.org/drawingml/2006/diagram}';A='{http://schemas.openxmlformats.org/drawingml/2006/main}';S='{http://schemas.microsoft.com/office/drawing/2008/diagram}'
output=[]
for row in json.loads(Path(sys.argv[1]).read_text()):
 file=Path(row['file']);assert hashlib.sha256(file.read_bytes()).hexdigest()==row['sha256']
 with zipfile.ZipFile(file) as z:
  assert z.testzip() is None
  root=E.fromstring(z.read(row['model']['part'].lstrip('/')));raw=list(root.iter());byindex={i+1:n for i,n in enumerate(raw)}
  points=list(root.find(D+'ptLst'));edges=list(root.find(D+'cxnLst'));ids=[p.get('modelId') for p in points+edges];assert len(ids)==len(set(ids))
  allpoints={p.get('modelId'):p for p in points};alledges={e.get('modelId'):e for e in edges}
  for point in points:
   if point.get('cxnId') is not None:assert point.get('cxnId') in alledges
   props=point.find(D+'prSet')
   if props is not None and props.get('presAssocID') is not None:assert props.get('presAssocID') in allpoints
  for edge in edges:
   assert edge.get('srcId') in allpoints and edge.get('destId') in allpoints
   for attribute,kind in [('parTransId','parTrans'),('sibTransId','sibTrans')]:
    if edge.get(attribute) is not None:
     transition=allpoints[edge.get(attribute)];assert transition.get('type')==kind and transition.get('cxnId')==edge.get('modelId')
  content=[p for p in points if p.get('type','node') in ['node','asst','doc']];byid={p.get('modelId'):p for p in content}
  docs=[p for p in content if p.get('type')=='doc'];assert len(docs)==1
  adjacent={p.get('modelId'):[] for p in content}
  for edge in edges:
   if edge.get('type','parOf')!='parOf':continue
   assert edge.get('srcId') in byid and edge.get('destId') in byid
   adjacent[edge.get('srcId')].append((int(edge.get('srcOrd')),edge.get('destId')))
  for siblings in adjacent.values():siblings.sort();assert len({order for order,key in siblings})==len(siblings)
  bindings={b['id']:byindex[b['elementIndex']] for b in row['projected']['bindings']};candidate=row['projected']['candidate']
  assert len(bindings)==len(content)-1
  assert [bindings[k].get('modelId') for k in candidate['roots']]==[key for order,key in adjacent[docs[0].get('modelId')]]
  for node in candidate['nodes']:
   source=bindings[node['id']];paragraphs=[]
   for para in source.findall('./'+D+'t/'+A+'p'):
    chunks=[]
    for run in para:
     if run.tag==A+'r':chunks.append(''.join(run.find(A+'t').itertext()))
     elif run.tag==A+'br':chunks.append('\n')
     elif run.tag==A+'fld':raise AssertionError('Unexpected field in owned plain-text experiment')
    paragraphs.append(''.join(chunks))
   assert paragraphs==node['paragraphs'];assert node['role']==('assistant' if source.get('type')=='asst' else 'content')
   assert [bindings[k].get('modelId') for k in node['children']]==[key for order,key in adjacent[source.get('modelId')]]
  allids={p.get('modelId') for p in points}
  shapes=[shape for n in z.namelist() if n.startswith('ppt/diagrams/drawing') and n.endswith('.xml') for shape in E.fromstring(z.read(n)).iter(S+'sp')]
  assert len(shapes)==row['cacheShapes'] and all(s.get('modelId') in allids for s in shapes)
  themes=[dict(part=n,name=E.fromstring(z.read(n)).get('name'),sha256=hashlib.sha256(z.read(n)).hexdigest()) for n in z.namelist() if n.startswith('ppt/theme/theme') and n.endswith('.xml')]
 output.append(dict(file=file.name,family=row['family'],sha256=row['sha256'],layout=row['layout'],nodes=len(candidate['nodes']),roots=len(candidate['roots']),assistants=sum(n['role']=='assistant' for n in candidate['nodes']),paragraphs=[p for n in candidate['nodes'] for p in n['paragraphs']],cacheShapes=len(shapes),themes=themes,exactNoOp=row['exactNoOp'],preview=row['preview'],independentTopologyTextRolesAndCache=True))
print(json.dumps(dict(python=sys.version.split()[0],records=output)))
`;
const independent = JSON.parse((await run("python3", ["-c", python, intermediate], { timeout: 60000, maxBuffer: 8 * 1024 * 1024 })).stdout);
const saveRounds = [];
for (const family of ["process", "hierarchy", "cycle"]) {
  const stage = observations.stages.filter(s => s.family === family).at(-1); let bytes = new Uint8Array(await readFile(path.join(directory, stage.file)));
  const target = path.join(directory, `${family}-unrelated-text-rounds`); await mkdir(target, { recursive: true });
  for (let round = 1; round <= 3; round++) {
    const session = await importPptx(bytes, { projection: "content", editSaveProfile: "text-runs-v3" });
    try {
      const slide = session.document.slides[0], current = slide.elements.find(e => e.type === "text"); assert.ok(current);
      const control = structuredClone(current); control.content[0].runs[0].text = `PPTKit native control ${family} round ${round}`;
      slide.removeElement(control.id); slide.insertElement(0, control);
      const prepared = await session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready");
      const saved = (await session.save(prepared.plan)).bytes, comparison = await compareOpcPackages(bytes, saved);
      assert.deepEqual(comparison.differences.map(d => d.part), ["/ppt/slides/slide1.xml"]);
      await writeFile(path.join(target, `round-${round}.pptx`), saved);
      saveRounds.push({ family, round, sha256: sha(saved), changedParts: ["/ppt/slides/slide1.xml"], unchangedDiagramPayloads: true, nativeReplay: "untested" }); bytes = saved;
    } finally { session.dispose(); }
  }
}
let identityControl;
if (observations.identityControl) {
  const control = observations.identityControl;
  assert.equal(sha(await readFile(control.source)), control.sha256);
  const identity = JSON.parse((await run("python3", ["-c", String.raw`
import sys,json,zipfile,xml.etree.ElementTree as E
from pathlib import Path
D='{http://schemas.openxmlformats.org/drawingml/2006/diagram}'
def inventory(file):
 with zipfile.ZipFile(file) as z:
  data=z.read('ppt/diagrams/data1.xml');root=E.fromstring(data)
  return data,{p.get('modelId') for p in root.find(D+'ptLst') if p.get('type','node') in ['node','asst']}
source,ids=inventory(sys.argv[1]);generated,_=inventory(sys.argv[2]);assert source==generated,'Native GUID control must exercise exactly the current writer data payload'
rows=[]
for file in json.loads(sys.argv[3]):
 _,current=inventory(file);assert current==ids
 rows.append(dict(file=Path(file).name,contentModelIds=sorted(current),contentIdentitySetRetained=True))
print(json.dumps(dict(contentModelIds=sorted(ids),sameDataAsCurrentGuidWriter=True,rounds=rows)))
`, control.source, control.generatedGuidCorpusFile, JSON.stringify(control.stages.map(file => path.join(directory, file)))], { timeout: 30000, maxBuffer: 1024 * 1024 })).stdout);
  identityControl = { ...control, ...identity, managedAuthorIdentityAdmission: "not-established" };
}
let deletionControl;
if (observations.deletionControl) {
  const control = observations.deletionControl;
  for (const name of [control.before, control.after]) {
    assert.equal(path.basename(name), name);
    assert.ok(records.some(r => r.file === path.join(directory, name)), "Deletion controls must be hash-verified stages.");
  }
  const checked = JSON.parse((await run("python3", ["-c", String.raw`
import sys,json,zipfile,xml.etree.ElementTree as E
D='{http://schemas.openxmlformats.org/drawingml/2006/diagram}';A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
def inventory(file):
 with zipfile.ZipFile(file) as z:
  root=E.fromstring(z.read('ppt/diagrams/data1.xml'));points=list(root.find(D+'ptLst'));edges=list(root.find(D+'cxnLst'))
  content={p.get('modelId'):(p.get('type','node'),[''.join(t.itertext()) for t in p.findall('./'+D+'t/'+A+'p')]) for p in points if p.get('type','node') in ['node','asst']}
  document=next(p.get('modelId') for p in points if p.get('type')=='doc')
  parents=[e for e in edges if e.get('type','parOf')=='parOf'];assert all(e.get('srcId')==document for e in parents)
  order=[e.get('destId') for e in sorted(parents,key=lambda e:int(e.get('srcOrd')))];assert len(order)==len(content)==len(set(order))
  return content,order
before,oldorder=inventory(sys.argv[1]);after,neworder=inventory(sys.argv[2]);removed=set(before)-set(after)
assert len(removed)==1 and not(set(after)-set(before));deleted=next(iter(removed));assert before[deleted]==('node',[''])
assert all(after[k]==before[k] for k in after);assert neworder==[k for k in oldorder if k!=deleted]
print(json.dumps(dict(deletedModelId=deleted,deletedParagraphs=before[deleted][1],beforeNodes=len(before),afterNodes=len(after),retainedContentModelIds=sorted(after),retainedTextRolesAndOrder=True)))
`, path.join(directory, control.before), path.join(directory, control.after)], { timeout: 30000, maxBuffer: 1024 * 1024 })).stdout);
  deletionControl = { ...control, ...checked, nativeClosureEvidence: "manual-reopened-screenshot-only" };
}
const evidence = { profile: "generated-smartart-native-edit-observations-v1", directory, application: observations.application,
  build: observations.build, os: observations.os, observations: observations.observations, additionalSourceObservation: observations.additionalSourceObservation,
  runtime: { node: process.version, python: independent.python }, inputs, artifacts, stages: independent.records, unrelatedTextSaveRounds: saveRounds,
  ...(identityControl ? { identityControl } : {}),
  ...(deletionControl ? { deletionControl } : {}),
  limitations: observations.limitations };
await writeFile(new URL("../docs/testing/smartart-generated-powerpoint-evidence.json", import.meta.url), JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ application: evidence.application, stages: evidence.stages.map(s => ({ file: s.file, nodes: s.nodes, assistants: s.assistants, cacheShapes: s.cacheShapes })), unrelatedTextSaveRounds: saveRounds.length }, null, 2));
