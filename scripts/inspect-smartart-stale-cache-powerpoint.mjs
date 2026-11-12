// Verifies retained local PowerPoint UI outputs; never drives/replays native UI.
import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { createHash } from "node:crypto";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { inspectPptx, importPptx } from "../packages/pptx-parser/dist/index.js";
const directory = path.resolve(process.argv[2] ?? ".tmp/smartart-stale-cache/powerpoint-2zdvkchy");
const sha = b => createHash("sha256").update(b).digest("hex"), run = promisify(execFile);
const planBytes = await readFile(path.join(directory, "plan.json")), plan = JSON.parse(planBytes);
const manifestBytes = await readFile(plan.sourceManifest); assert.equal(sha(manifestBytes), plan.sourceManifestSha256);
const manifest = JSON.parse(manifestBytes), ui = JSON.parse(await readFile(path.join(directory, "ui-observations.json"))), relayout = JSON.parse(await readFile(path.join(directory, "relayout-observations.json")));
const rounds = JSON.parse(await readFile(path.join(directory, "rounds-plan.json"))), roundsUi = JSON.parse(await readFile(path.join(directory, "rounds-observations.json")));
assert.equal(ui.cases.length, 15); assert.equal(relayout.cases.length, 3); assert.equal(rounds.length, 6); assert.equal(roundsUi.cases.length, 6);
const inputs = [...plan.cases.map(c => ({ ...c, name: `${c.family}-${c.kind}` })), ...relayout.cases.map(c => ({ ...plan.cases.find(p => p.family === c.family && p.kind === "data-only"), kind: "relayout", name: `${c.family}-relayout`, operations: c.operations })), ...rounds];
const records = [];
for (const input of inputs) {
  const source = input.source ? { file: input.source, sha256: input.sha256 } : manifest.cases.find(c => c.family === input.family && c.kind === (input.kind === "relayout" ? "data-only" : input.kind));
  assert.equal(sha(await readFile(source.file)), input.sha256, "Immutable control changed");
  const file = path.join(directory, input.name + "-saved.pptx"), bytes = await readFile(file);
  assert.deepEqual(bytes, await readFile(path.join(directory, input.name + ".pptx")), "Reopened copy differs from saved snapshot");
  const report = await inspectPptx(bytes, { diagramTextComparison: "direct-paragraphs" });
  const inventory = report.diagramInventory; assert.equal(inventory.dataModels.length, 1); assert.deepEqual(inventory.dataModels[0].findings, []);
  const pairs = inventory.textComparisons; assert.equal(pairs.length, 1); assert.equal(pairs[0].freshness, "unknown");
  assert.ok(pairs[0].shapes.every(s => s.status !== "different"), "Native saved source/cache direct text differs");
  const selected = pairs[0].shapes.filter(s => s.sourceParagraphs?.includes(input.marker));
  const expectMarker = ["data-only", "both", "relayout", "round2", "round3"].includes(input.kind);
  if (expectMarker) assert.ok(selected.length && selected.every(s => s.status === "equal"));
  else assert.equal(selected.length, 0);
  const cacheShapes = inventory.cacheMappings[0].shapes.length; assert.ok(cacheShapes > 0, "Native cache remains empty");
  const session = await importPptx(bytes, { projection: "content", diagramPreview: "cache", diagramTextComparison: "direct-paragraphs" });
  let preview;
  try { const p = await session.prepareSave({ mode: "no-op" }); assert.equal(p.status, "ready"); assert.deepEqual((await session.save(p.plan)).bytes, new Uint8Array(bytes)); preview = session.report.diagramPreviews; }
  finally { session.dispose(); }
  const artifacts = [];
  for (const stage of input.kind === "relayout" ? ["open", "intermediate", "applied", "reopened"] : ["open", "reopened"]) for (const ext of ["png", "ax.txt"]) {
    const artifact = `${input.name}-${stage}.${ext}`, data = await readFile(path.join(directory, artifact));
    assert.ok(data.length); artifacts.push({ file: artifact, sha256: sha(data), byteLength: data.length });
  }
  records.push({ family: input.family, kind: input.kind, file, sha256: sha(bytes), source: source.file, sourceSha256: source.sha256, marker: input.marker, selection: input.selection,
    layout: report.slides[0].diagrams[0].components.find(c => c.role === "layout").layoutIdentifier, ...(input.operations ? { operations: input.operations } : {}), ...(input.previousSaved ? { previousSaved: input.previousSaved, previousSavedSha256: input.previousSavedSha256 } : {}),
    comparison: pairs[0], cacheShapes, preview, exactNoOp: true, artifacts, ui: input.kind === "relayout" ? relayout.observed : input.kind.startsWith("round") ? roundsUi.cases.find(c => c.family === input.family && c.kind === input.kind) : ui.cases.find(c => c.family === input.family && c.kind === input.kind) });
}
const intermediate = path.join(directory, "native-text-inspection.json");
await writeFile(intermediate, JSON.stringify(records, null, 2) + "\n");
const independent = JSON.parse((await run("python3", ["-c", String.raw`
import sys,json,hashlib,zipfile,uuid,xml.etree.ElementTree as E
from pathlib import Path
D='{http://schemas.openxmlformats.org/drawingml/2006/diagram}';A='{http://schemas.openxmlformats.org/drawingml/2006/main}';S='{http://schemas.microsoft.com/office/drawing/2008/diagram}'
def key(v):
 try:return str(uuid.UUID(v))
 except ValueError:
  try:return str(int(v))
  except ValueError:return v
def paragraphs(body):
 if body is None:return None
 return [''.join(('\n' if n.tag==A+'br' else ''.join(t.text or '' for t in n.findall(A+'t'))) for n in p if n.tag in [A+'r',A+'br']) for p in body.findall(A+'p')]
def inspect(file,expected):
 assert hashlib.sha256(Path(file).read_bytes()).hexdigest()==expected
 with zipfile.ZipFile(file) as z:
  assert z.testzip() is None
  models=[];caches=[];payloads={n:hashlib.sha256(z.read(n)).hexdigest() for n in z.namelist()}
  for name in z.namelist():
   if name.startswith('ppt/diagrams/') and name.endswith('.xml'):
    root=E.fromstring(z.read(name))
    if root.tag==D+'dataModel':models.append((name,root))
    if root.tag==S+'drawing':caches.append((name,root))
  assert len(models)==len(caches)==1
  name,model=models[0];points={key(p.get('modelId')):p for p in model.find(D+'ptLst')};assert len(points)==len(model.find(D+'ptLst'))
  graph={i:dict(kind=p.get('type','node'),paragraphs=paragraphs(p.find(D+'t'))) for i,p in points.items() if p.get('type','node') in ['node','asst','doc']}
  edges=[{k:(key(e.get(k)) if k in ['modelId','srcId','destId'] else e.get(k)) for k in ['modelId','srcId','destId','srcOrd','destOrd']} for e in model.find(D+'cxnLst') if e.get('type','parOf')=='parOf' and key(e.get('srcId')) in graph and key(e.get('destId')) in graph]
  shapes=[]
  for s in caches[0][1].iter(S+'sp'):
   target=key(s.get('modelId'));point=points.get(target);evidence=[]
   if point is not None and point.get('type')=='pres':
    prop=point.find(D+'prSet')
    if prop is not None and prop.get('presAssocID') is not None:evidence.append(key(prop.get('presAssocID')))
    evidence.extend(key(c.get('srcId')) for c in model.find(D+'cxnLst') if c.get('type')=='presOf' and key(c.get('destId'))==target)
    target=evidence[0] if evidence and len(set(evidence))==1 else None
   texts=paragraphs(s.find(S+'txBody'));source=graph.get(target)
   equal=None if source is None or source['paragraphs'] is None or texts is None else texts==source['paragraphs']
   shapes.append(dict(modelId=s.get('modelId'),target=target,paragraphs=texts,equal=equal))
  return dict(dataPart='/'+name,cachePart='/'+caches[0][0],graph=graph,edges=sorted(edges,key=lambda e:e['modelId']),shapes=shapes,payloads=payloads,statistics=dict(zipBytes=Path(file).stat().st_size,parts=len(payloads),relationships=sum(len(E.fromstring(z.read(n))) for n in z.namelist() if n.endswith('.rels')),mediaParts=sum(n.startswith('ppt/media/') for n in z.namelist()),diagramParts=sum(n.startswith('ppt/diagrams/') for n in z.namelist())))
rows=[]
for row in json.loads(Path(sys.argv[1]).read_text()):
 before=inspect(row['source'],row['sourceSha256']);after=inspect(row['file'],row['sha256'])
 assert all(s['equal'] is not False for s in after['shapes']),(row['family'],row['kind'],'saved direct text differs')
 assert before['graph']==after['graph'],(row['family'],row['kind'],'content identity/role/text changed')
 assert before['edges']==after['edges'],(row['family'],row['kind'],'content connections/order changed')
 assert len(after['shapes'])==row['cacheShapes']
 expected=row['kind'] in ['data-only','both','relayout','round2','round3'];marker=row['marker']
 assert any(marker in p for n in after['graph'].values() for p in n['paragraphs'] or [])==expected
 assert any(marker in p for n in after['shapes'] for p in n['paragraphs'] or [])==expected
 if row['kind'].startswith('round'):
  previous=inspect(row['previousSaved'],row['previousSavedSha256'])
  assert set(previous['payloads'])==set(before['payloads'])
  assert [n for n in previous['payloads'] if previous['payloads'][n]!=before['payloads'][n]]==[before['dataPart'].lstrip('/')]
  assert previous['edges']==before['edges']
  assert set(previous['graph'])==set(before['graph'])
  changed=[i for i in previous['graph'] if previous['graph'][i]!=before['graph'][i]]
  assert changed==[key(row['selection']['sourceModelId'])]
  assert any(s['equal'] is False for s in before['shapes'])
  assert previous['statistics']['parts']==after['statistics']['parts']
  assert previous['statistics']['relationships']==after['statistics']['relationships']
  assert previous['statistics']['mediaParts']==after['statistics']['mediaParts']
 allparts=set(before['payloads'])|set(after['payloads']);changed=sorted(n for n in allparts if before['payloads'].get(n)!=after['payloads'].get(n))
 rows.append(dict(family=row['family'],kind=row['kind'],contentIdentityRolesTextAndOrderRetained=True,dataPart=after['dataPart'],cachePart=after['cachePart'],contentPoints=len(after['graph'])-1,assistants=sum(n['kind']=='asst' for n in after['graph'].values()),cacheShapes=len(after['shapes']),comparableShapes=sum(s['equal'] is not None for s in after['shapes']),changedParts=['/'+n for n in changed],sourceGraph=before['graph'],nativeGraph=after['graph'],nativeCache=after['shapes'],sourceStatistics=before['statistics'],nativeStatistics=after['statistics']))
print(json.dumps(rows))
`, intermediate], { timeout: 30000, maxBuffer: 8 * 1024 * 1024 }).catch(error => { throw new Error(error.stderr || error.message); })).stdout);
const result = { profile: "powerpoint-stale-cache-observation-v1", directory, application: ui.application, version: ui.version, build: ui.build, sourceManifestSha256: plan.sourceManifestSha256, planSha256: sha(planBytes),
  records, independent, limitations: ["Retained UI operations are separately observed, not replayed by this verifier.", "These twenty-four named standard-layout cases do not admit arbitrary SmartArt editing, cache freshness or geometry/style/font fidelity.", "Long marker labels trigger native text fitting; visuals differ from baseline.", "Preservation no-op is verified; SmartArt internal edit-save remains unsupported.", "Native/source artifacts remain local pending redistribution admission."] };
await writeFile(process.argv[3] ? path.resolve(process.argv[3]) : new URL("../docs/testing/smartart-stale-cache-powerpoint-evidence.json", import.meta.url), JSON.stringify(result, null, 2) + "\n");
console.log(JSON.stringify({ records: records.length, independent: independent.length, directory }, null, 2));
