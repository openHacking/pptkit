// Optional controlled stale-cache corpus. Original files remain unchanged.
// ElementTree rewrites selected test parts; this is not a preservation-save API.
import assert from "node:assert/strict";
import { readFile, writeFile, mkdir, mkdtemp } from "node:fs/promises";
import { createHash } from "node:crypto";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { inspectPptx, importPptx } from "../packages/pptx-parser/dist/index.js";
const run = promisify(execFile), sha = data => createHash("sha256").update(data).digest("hex");
const manifestPath = path.resolve(process.argv[2] ?? "docs/testing/smartart-generated-powerpoint-evidence.json"), originalManifest = await readFile(manifestPath), manifest = JSON.parse(originalManifest);
const executable = process.argv[3], nativeEnabled = !!executable;
await mkdir(".tmp/smartart-stale-cache", { recursive: true }); const directory = await mkdtemp(path.resolve(".tmp/smartart-stale-cache/run-"));
const config = { diagramTextComparison: "direct-paragraphs" }, cases = [];
const plainText = content => typeof content === "string" ? content : content.map(p => p.runs.map(r => r.text).join("")).join("\n");
const textOf = elements => elements.flatMap(e => e.type === "group" ? textOf(e.children) : e.type === "text" ? [plainText(e.content)] : e.type === "shape" && e.text ? [plainText(e.text.content)] : []);
for (const [family, file] of [["process", "process-round2-theme.pptx"], ["hierarchy", "hierarchy-round2-theme.pptx"], ["cycle", "cycle-guid-round3.pptx"]]) {
  const stage = manifest.stages.find(s => s.file === file); assert.ok(stage);
  const source = path.join(manifest.directory, file), original = await readFile(source); assert.equal(sha(original), stage.sha256);
  const inspection = await inspectPptx(original, config), comparison = inspection.diagramInventory.textComparisons[0];
  const eligible = comparison.shapes.find(s => s.status === "equal" && s.nonemptyText && s.sourceParagraphs.length === 1 && !s.sourceParagraphs[0].includes("\n")); assert.ok(eligible);
  const sourceModel = inspection.diagramInventory.dataModels.find(m => m.part === comparison.dataPart), selectedPoint = sourceModel.points.find(p => p.elementIndex === eligible.sourcePointElementIndex);
  const selection = { dataPart: comparison.dataPart, cachePart: comparison.cachePart, pointIndex: eligible.sourcePointElementIndex, sourceModelId: selectedPoint.modelId,
    cacheIndices: comparison.shapes.filter(s => s.status === "equal" && s.nonemptyText && s.sourcePointElementIndex === eligible.sourcePointElementIndex).map(s => s.cacheElementIndex), originalParagraphs: eligible.sourceParagraphs };
  for (const kind of ["baseline", "data-only", "cache-only", "both", "empty"]) {
    const output = path.join(directory, `${family}-${kind}.pptx`), marker = `PPTKit${family}${kind.replaceAll("-", "")}Control`;
    const transformed = JSON.parse((await run("python3", ["-c", String.raw`
import sys,json,hashlib,zipfile,xml.etree.ElementTree as E
from pathlib import Path
source,target,kind,marker,expected=sys.argv[1:6];selection=json.loads(sys.argv[6]);assert hashlib.sha256(Path(source).read_bytes()).hexdigest()==expected
D='{http://schemas.openxmlformats.org/drawingml/2006/diagram}';A='{http://schemas.openxmlformats.org/drawingml/2006/main}';S='{http://schemas.microsoft.com/office/drawing/2008/diagram}'
with zipfile.ZipFile(source) as z:
 assert z.testzip() is None
 parts={n:z.read(n) for n in z.namelist()};before={n:hashlib.sha256(b).hexdigest() for n,b in parts.items()}
 data=E.fromstring(parts[selection['dataPart'].lstrip('/')]);cache=E.fromstring(parts[selection['cachePart'].lstrip('/')]);draw=list(data.iter());cached=list(cache.iter())
 point=draw[selection['pointIndex']-1];assert point.tag==D+'pt' and point.get('modelId')==selection['sourceModelId']
 def patch(owner,path):
  leaves=owner.findall(path);assert leaves
  for i,leaf in enumerate(leaves):leaf.text=marker if i==0 else ''
 if kind in ['data-only','both']:
  patch(point,'./'+D+'t/'+A+'p/'+A+'r/'+A+'t');parts[selection['dataPart'].lstrip('/')]=E.tostring(data,encoding='utf-8',xml_declaration=True)
 if kind in ['cache-only','both']:
  for i in selection['cacheIndices']:
   shape=cached[i-1];assert shape.tag==S+'sp';patch(shape,'./'+S+'txBody/'+A+'p/'+A+'r/'+A+'t')
  parts[selection['cachePart'].lstrip('/')]=E.tostring(cache,encoding='utf-8',xml_declaration=True)
 if kind=='empty':
  tree=cache.find(S+'spTree');assert tree is not None
  for child in list(tree):
   if child.tag not in [S+'nvGrpSpPr',S+'grpSpPr']:tree.remove(child)
  parts[selection['cachePart'].lstrip('/')]=E.tostring(cache,encoding='utf-8',xml_declaration=True)
 changed=sorted(n for n,b in parts.items() if hashlib.sha256(b).hexdigest()!=before[n]);expected_parts=[]
 if kind in ['data-only','both']:expected_parts.append(selection['dataPart'].lstrip('/'))
 if kind in ['cache-only','both','empty']:expected_parts.append(selection['cachePart'].lstrip('/'))
 assert changed==sorted(expected_parts)
 if kind=='baseline':Path(target).write_bytes(Path(source).read_bytes())
 else:
  with zipfile.ZipFile(target,'w',compression=zipfile.ZIP_DEFLATED) as z:
   for n,b in parts.items():z.writestr(n,b)
with zipfile.ZipFile(target) as z:
 assert z.testzip() is None and set(z.namelist())==set(before)
 assert all(hashlib.sha256(z.read(n)).hexdigest()==before[n] for n in before if n not in changed)
print(json.dumps(dict(sha256=hashlib.sha256(Path(target).read_bytes()).hexdigest(),changedParts=['/'+n for n in changed],unmodifiedPayloadsExact=True)))
`, source, output, kind, marker, stage.sha256, JSON.stringify(selection)], { timeout: 30000, maxBuffer: 2 * 1024 * 1024 })).stdout);
    const bytes = await readFile(output), checked = await inspectPptx(bytes, config), pair = checked.diagramInventory.textComparisons[0];
    const selected = pair.shapes.filter(s => selection.cacheIndices.includes(s.cacheElementIndex));
    if (kind !== "empty") {
      assert.equal(selected.length, selection.cacheIndices.length);
      assert.ok(selected.every(s => s.status === (["data-only", "cache-only"].includes(kind) ? "different" : "equal")));
      assert.deepEqual(selected[0].sourceParagraphs, ["data-only", "both"].includes(kind) ? [marker] : selection.originalParagraphs);
      assert.deepEqual(selected[0].cacheParagraphs, ["cache-only", "both"].includes(kind) ? [marker] : selection.originalParagraphs);
    } else assert.deepEqual(pair.shapes, []);
    const session = await importPptx(bytes, { projection: "content", diagramPreview: "cache", ...config });
    let preview;
    try {
      const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
      const cachedText = textOf(session.document.slides[0].elements.filter(e => e.type === "group"));
      preview = { report: session.report.diagramPreviews, cachedText, markerPresent: cachedText.some(t => t.includes(marker)) };
      if (["cache-only", "both"].includes(kind)) assert.equal(preview.markerPresent, true);
      if (["data-only", "baseline", "empty"].includes(kind)) assert.equal(preview.markerPresent, false);
      assert.ok(preview.report.every(p => p.freshness === "unknown"));
    } finally { session.dispose(); }
    assert.equal(sha(await readFile(source)), stage.sha256);
    cases.push({ family, kind, file: output, marker, source, sourceSha256: stage.sha256, selection, ...transformed, comparison: pair, selected, preview, exactNoOp: true });
  }
}
await writeFile(path.join(directory, "controls.json"), JSON.stringify(cases, null, 2) + "\n");
let application, native = [];
if (nativeEnabled) {
  const options = { timeout: 45000, maxBuffer: 8 * 1024 * 1024 };
  application = (await run(executable, ["--version"], options)).stdout.trim();
  const direct = path.join(directory, "direct"), saved = path.join(directory, "saved"), reopened = path.join(directory, "reopened");
  await mkdir(direct); await mkdir(saved); await mkdir(reopened);
  const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, "profile")).href}`;
  for (const sample of cases) {
    const log = { family: sample.family, kind: sample.kind };
    try {
      log.direct = await run(executable, [profile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", direct, sample.file], options);
      log.save = await run(executable, [profile, "--headless", "--convert-to", "pptx:Impress Office Open XML", "--outdir", saved, sample.file], options);
      const output = path.join(saved, path.basename(sample.file));
      log.reopen = await run(executable, [profile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", reopened, output], options);
      const observed = JSON.parse((await run("python3", ["-c", String.raw`
import sys,json,hashlib,zipfile,xml.etree.ElementTree as E,fitz
from pathlib import Path
sample=json.loads(sys.argv[1]);source=Path(sample['file']);saved=Path(sys.argv[2]);direct=Path(sys.argv[3]);reopened=Path(sys.argv[4]);assert hashlib.sha256(source.read_bytes()).hexdigest()==sample['sha256']
D='{http://schemas.openxmlformats.org/drawingml/2006/diagram}';A='{http://schemas.openxmlformats.org/drawingml/2006/main}';S='{http://schemas.microsoft.com/office/drawing/2008/diagram}'
def text(file):
 with fitz.open(file) as doc:return '\n'.join(p.get_text() for p in doc)
def model_text(file):
 out=[];cache=[]
 with zipfile.ZipFile(file) as z:
  assert z.testzip() is None
  for name in z.namelist():
   if not(name.startswith('ppt/diagrams/') and name.endswith('.xml')):continue
   root=E.fromstring(z.read(name))
   if root.tag==D+'dataModel':
    for p in root.find(D+'ptLst'):
     if p.get('type','node') in ['node','asst']:out.extend(''.join(n.text or '' for n in p.iter(A+'t')) for _ in [0])
   if root.tag==S+'drawing':cache.extend(''.join(n.text or '' for n in p.iter(A+'t')) for p in root.iter(S+'sp'))
 return out,cache
source_text,source_cache=model_text(source);saved_text,saved_cache=model_text(saved);a=text(direct);b=text(reopened);marker=sample['marker']
match=lambda s,v:''.join(v.split()) in ''.join(s.split())
print(json.dumps(dict(status='saved-and-reopened',nativeSha256=hashlib.sha256(saved.read_bytes()).hexdigest(),directPdfSha256=hashlib.sha256(direct.read_bytes()).hexdigest(),reopenedPdfSha256=hashlib.sha256(reopened.read_bytes()).hexdigest(),directText=a,reopenedText=b,directMarkerPresent=match(a,marker),reopenedMarkerPresent=match(b,marker),directOriginalLabelPresent=match(a,sample['selection']['originalParagraphs'][0]),reopenedOriginalLabelPresent=match(b,sample['selection']['originalParagraphs'][0]),sourceText=source_text,sourceCacheText=source_cache,nativeText=saved_text,nativeCacheText=saved_cache,independentZipXmlPdf=True)))
`, JSON.stringify(sample), output, path.join(direct, path.basename(sample.file, ".pptx") + ".pdf"), path.join(reopened, path.basename(sample.file, ".pptx") + ".pdf")], options)).stdout);
      native.push({ family: sample.family, kind: sample.kind, ...observed, log });
    } catch (error) { native.push({ family: sample.family, kind: sample.kind, status: "native-operation-failed", message: error.message, log }); }
    await writeFile(path.join(directory, "native-progress.json"), JSON.stringify(native, null, 2) + "\n");
    console.log(JSON.stringify({ family: sample.family, kind: sample.kind, status: native.at(-1).status }));
  }
}
const result = { profile: "smartart-stale-cache-controls-v1", directory, nativeEnabled, ...(application ? { application } : {}), originApplication: manifest.application, originBuild: manifest.build, manifestSha256: sha(originalManifest), cases, native,
  limitations: ["Controlled test-part transformations only; not a source-preservation write API.", "Different direct paragraph vectors do not establish cache staleness or text ownership for every native layout; equal vectors do not establish freshness.", "Cached preview deliberately follows stored cache characters and keeps unknown freshness.", "Native observations, when enabled, are named headless open/save/PDF reopen only; PowerPoint stale-cache UI behavior and arbitrary compatibility remain untested.", "Source/native/transformed third-party artifacts remain local pending fixture redistribution admission."] };
await writeFile(new URL("../docs/testing/smartart-stale-cache-evidence.json", import.meta.url), JSON.stringify(result, null, 2) + "\n");
console.log(JSON.stringify({ directory, controls: cases.length, native: native.length, application }, null, 2));
