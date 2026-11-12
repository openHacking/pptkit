// Optional native producer experiment; portable CI does not require soffice.
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdtemp, mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { createHash } from "node:crypto";
import { platform, release, arch } from "node:os";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { createContentImportFixture } from "../tests/fixtures/content-import.mjs";

if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-content-import.mjs [soffice-executable]");
const run = promisify(execFile), executable = process.argv[2] ?? "soffice", options = { timeout: 45000, maxBuffer: 4 * 1024 * 1024 };
const version = (await run(executable, ["--version"], options)).stdout.trim();
await mkdir(".tmp/content-import", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/content-import/run-")), savedDirectory = path.join(directory, "native"), reopenedDirectory = path.join(directory, "reopened");
await mkdir(savedDirectory); await mkdir(reopenedDirectory);
const original = await generatePptx(createContentImportFixture(createPresentation)), filename = path.join(directory, "source.pptx");
await writeFile(filename, original.bytes);
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, "profile")).href}`;
const nativeLog = await run(executable, [profile, "--headless", "--convert-to", "pptx:Impress Office Open XML", "--outdir", savedDirectory, filename], options);
const nativeFilename = path.join(savedDirectory, "source.pptx"), nativeBytes = await readFile(nativeFilename);
const session = await importPptx(nativeBytes, { projection: "content" });
try {
  const prepared = await session.prepareSave({ mode: "no-op" }); assert.equal(prepared.status, "ready");
  const noOp = await session.save(prepared.plan); assert.equal(Buffer.compare(Buffer.from(noOp.bytes), nativeBytes), 0); prepared.plan.dispose();
  await writeFile(path.join(directory, "import-report.json"), JSON.stringify(session.report, null, 2) + "\n");
  const projected = normalizePresentation(session.document);
  const coreTexts = [];
  const visit = elements => { for (const element of elements) {
    if (element.type === "group") visit(element.children);
    else if (element.type === "text") coreTexts.push(element.plainText);
    else if (element.type === "shape" && element.text) coreTexts.push(element.text.plainText);
  } };
  for (const slide of projected.slides) visit(slide.elements);
  // Fresh export is deliberately separate from the source-preservation no-op.
  const fresh = await generatePptx(session.document), freshFilename = path.join(directory, "projection.pptx");
  await writeFile(freshFilename, fresh.bytes);
  const reopenLog = await run(executable, [profile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", reopenedDirectory, freshFilename], options);
  const pdf = await readFile(path.join(reopenedDirectory, "projection.pdf")); assert.ok(pdf.subarray(0, 5).equals(Buffer.from("%PDF-")));
  const inventory = JSON.parse((await run("python3", ["-c", `
import sys, zipfile, json, xml.etree.ElementTree as E
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
result=[]
for filename in sys.argv[1:]:
    texts=[]
    with zipfile.ZipFile(filename) as z:
        assert z.testzip() is None
        for part in sorted(n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')):
            root=E.fromstring(z.read(part))
            for body in root.findall('.//p:txBody',ns):
                text='\\n'.join(''.join(n.text or '' for n in p.findall('.//a:t',ns)) for p in body.findall('a:p',ns))
                if text: texts.append(text)
    result.append(texts)
print(json.dumps(result))
`, nativeFilename, freshFilename], options)).stdout);
  const compare = (a, b) => JSON.stringify(a.filter(Boolean).toSorted()) === JSON.stringify(b.filter(Boolean).toSorted());
  const hash = bytes => createHash("sha256").update(bytes).digest("hex");
  const features = {};
  for (const finding of session.report.findings) if (finding.operation === "import") {
    const key = `${finding.feature}:${finding.tier}`; features[key] = (features[key] ?? 0) + 1;
  }
  const evidence = {
    schemaVersion: 1, application: version, platform: `${platform()} ${release()} ${arch()}`,
    artifactsDirectory: path.relative(process.cwd(), directory),
    scope: "LibreOffice-produced source content projection, exact no-op bytes and fresh projection reopen; not preservation edit-save, visual parity, interactive editability or PowerPoint certification",
    original: { sha256: hash(original.bytes) }, nativeSource: { sha256: hash(nativeBytes), textInventory: inventory[0] },
    projection: { coreTexts, mappedElements: session.report.identities.filter(i => i.kind === "element").length, importFeatures: features },
    noOpOriginalBytes: true,
    nativeTextInventoryMapped: compare(inventory[0], coreTexts),
    freshExport: { sha256: hash(fresh.bytes), textInventory: inventory[1], sourceTextInventoryRetained: compare(inventory[0], inventory[1]) },
    reopenedPdf: { sha256: hash(pdf), byteLength: pdf.length },
    logs: { native: nativeLog, reopen: reopenLog },
  };
  await writeFile("docs/testing/content-import-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
  console.log(JSON.stringify({ directory, nativeTextInventoryMapped: evidence.nativeTextInventoryMapped, freshTextInventoryRetained: evidence.freshExport.sourceTextInventoryRetained, mappedElements: evidence.projection.mappedElements }));
} finally { session.dispose(); }
