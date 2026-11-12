// Optional native-layout baseline, not a public importer or layout engine.
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { readFile, writeFile, mkdir, mkdtemp } from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { inspectPptx } from "../packages/pptx-parser/dist/index.js";
import { observeDiagramGeometry, compareNativeGeometryObservations } from "./research/diagram-geometry-observation.mjs";
import { pdfTextEvidence } from "./research/pdf-text-evidence.mjs";

const prepared = JSON.parse(await readFile(new URL("../docs/testing/smartart-native-presentation-candidate-evidence.json", import.meta.url), "utf8"));
const manifestFile = path.resolve(process.argv[2] ?? path.join(prepared.directory, "manifest.json"));
const sourceDirectory = path.dirname(manifestFile), manifest = JSON.parse(await readFile(manifestFile, "utf8"));
assert.equal(manifest.records.length, 24, "The baseline requires the complete 24-case candidate corpus.");
for (const family of ["process", "hierarchy", "cycle"]) for (const kind of ["empty", "single", "standard", "long"]) for (const cache of ["absent", "empty"]) {
  assert.equal(manifest.records.filter(row => row.family === family && row.case === kind && row.cache === cache && row.presentation === "layout").length, 1, "Each family/case/cache combination must be registered exactly once.");
}
const executable = process.argv[3] ?? "soffice", run = promisify(execFile), sha = bytes => createHash("sha256").update(bytes).digest("hex");
const directory = await mkdtemp(path.join(sourceDirectory, "geometry-run-")), fodp = path.join(directory, "fodp"), pdf = path.join(directory, "pdf");
await mkdir(fodp); await mkdir(pdf);
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, "profile")).href}`;
const commandOptions = { timeout: 45000, maxBuffer: 16 * 1024 * 1024 };
const application = (await run(executable, ["--version"], commandOptions)).stdout.trim();
const extractor = fileURLToPath(new URL("./research/extract-native-diagram-geometry.py", import.meta.url));
const records = [];
for (const row of manifest.records) {
  assert.equal(path.basename(row.file), row.file); assert.equal(row.presentation, "layout");
  const input = path.join(sourceDirectory, row.file), before = await readFile(input); assert.equal(sha(before), row.sha256);
  const record = { file: row.file, family: row.family, case: row.case, cache: row.cache, inputSha256: row.sha256 };
  try {
    const inspection = await inspectPptx(before); assert.equal(inspection.diagramInventory.dataModels.length, 1);
    const fodpLog = await run(executable, [profile, "--headless", "--convert-to", "fodp", "--outdir", fodp, input], commandOptions);
    const pdfLog = await run(executable, [profile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", pdf, input], commandOptions);
    const nativeFile = path.join(fodp, row.file.replace(/\.pptx$/, ".fodp")), rendered = path.join(pdf, row.file.replace(/\.pptx$/, ".pdf"));
    const extracted = JSON.parse((await run("python3", [extractor, input, nativeFile], commandOptions)).stdout);
    const geometry = observeDiagramGeometry(extracted.snapshot, inspection.diagramInventory.dataModels[0]);
    assert.equal(geometry.shapes.length, extracted.oracle.length);
    for (const [i, expected] of extracted.oracle.entries()) {
      const actual = geometry.shapes[i]; assert.equal(actual.occurrence, expected.ordinal); assert.equal(actual.geometryStatus, expected.status);
      if (expected.boundsPoints) for (const key of ["x", "y", "width", "height"]) assert.ok(Math.abs(actual.boundsPoints[key] - expected.boundsPoints[key]) < 1e-8);
    }
    const pdfObservation = JSON.parse((await run("python3", ["-c", String.raw`
import sys,json,fitz,hashlib
from pathlib import Path
p=Path(sys.argv[1]);d=fitz.open(p);pages=[]
for i,page in enumerate(d):
 spans=[dict(text=s['text'],boundsPoints=list(s['bbox']),font=s['font'],fontSizePoints=s['size']) for b in page.get_text('dict')['blocks'] if b['type']==0 for line in b['lines'] for s in line['spans']]
 paths=[dict(boundsPoints=list(p['rect']),paint=p['type']) for p in page.get_drawings()]
 if len(spans)>10000 or len(paths)>10000:raise ValueError('PDF observation budget exceeded')
 picture=p.with_suffix('.png');page.get_pixmap(matrix=fitz.Matrix(1,1)).save(str(picture))
 pages.append(dict(page=i+1,boundsPoints=list(page.rect),text=page.get_text(),spans=spans,paths=paths,picture=str(picture),pictureSha256=hashlib.sha256(picture.read_bytes()).hexdigest()))
print(json.dumps(dict(pages=pages,pymupdf=fitz.VersionBind)))
`, rendered], commandOptions)).stdout);
    const markers = row.authored.nodes.flatMap(n => n.paragraphs).filter(p => p.trim());
    const drawingText = geometry.shapes.flatMap(shape => shape.raw.paragraphs).join("\n");
    Object.assign(record, { status: "observed", groupAssociation: extracted.groupAssociation, geometry, fontFaces: extracted.fontFaces,
      pdf: pdfObservation, text: { markerCount: markers.length, ...pdfTextEvidence(pdfObservation.pages.map(p => p.text).join("\n"), markers) },
      storedDrawingText: { markerCount: markers.length, ...pdfTextEvidence(drawingText, markers) },
      nativeSha256: sha(await readFile(nativeFile)), pdfSha256: sha(await readFile(rendered)), independentGeometryChecked: true, logs: { fodp: fodpLog, pdf: pdfLog } });
  } catch (error) {
    Object.assign(record, { status: "observation-failed", failure: { message: error.message, code: error.code, stdout: error.stdout, stderr: error.stderr } });
  }
  assert.equal(sha(await readFile(input)), row.sha256, "Native observation must not change the input candidate.");
  records.push(record); console.log(JSON.stringify({ file: row.file, status: record.status, group: record.groupAssociation?.status, shapes: record.geometry?.summary, text: record.text?.textMatchStatus }));
}
// Controlled cache-policy comparison only. Producer-generated unmatched names
// (e.g. a random background UUID) are intentionally not geometry identities.
const cachePolicyComparisons = [];
for (const family of ["process", "hierarchy", "cycle"]) for (const kind of ["empty", "single", "standard", "long"]) {
  const pair = ["absent", "empty"].map(cache => records.find(r => r.family === family && r.case === kind && r.cache === cache));
  const comparable = pair.every(r => r?.status === "observed" && r.groupAssociation.status === "located");
  cachePolicyComparisons.push({ family, case: kind, inputs: pair.map(r => r?.inputSha256),
    status: comparable ? "compared" : "not-comparable",
    ...(comparable ? compareNativeGeometryObservations(pair[0].geometry, pair[1].geometry) : {}) });
}
const evidence = { profile: "native-diagram-geometry-baseline-v1", application, directory, sourceDirectory, manifestSha256: sha(await readFile(manifestFile)), records, cachePolicyComparisons,
  limitations: ["Fresh direct FODP/PDF producer observations, not PowerPoint interactive editing or a native compatibility certificate.", "ODF shape names may repeat and do not establish stable semantic node identity; association evidence is explicit and occurrence-scoped.", "Transformed boxes/text areas/clipping remain source-only; PDF spans/paint bounds are separate observations without fabricated shape ownership.", "Same-run artifacts do not certify cross-platform fonts, deterministic native geometry or PPTKit layout parity.", "Third-party resources and rendered artifacts stay local pending redistribution/fixture licensing admission."] };
await writeFile(new URL("../docs/testing/smartart-native-geometry-evidence.json", import.meta.url), JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ directory, cases: records.length, observed: records.filter(r => r.status === "observed").length }));
