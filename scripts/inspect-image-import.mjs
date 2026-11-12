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
import { createImageImportFixture } from "../tests/fixtures/content-import.mjs";
import { png } from "../packages/pptx-parser/test/image-fixture.mjs";

if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-image-import.mjs [soffice-executable]");
const run = promisify(execFile), executable = process.argv[2] ?? "soffice", options = { timeout: 45000, maxBuffer: 4 * 1024 * 1024 };
const version = (await run(executable, ["--version"], options)).stdout.trim();
await mkdir(".tmp/image-import-native", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/image-import-native/run-"));
const nativeDirectory = path.join(directory, "native"), reopenedDirectory = path.join(directory, "reopened");
await mkdir(nativeDirectory); await mkdir(reopenedDirectory);
const original = await generatePptx(createImageImportFixture(createPresentation, png)), filename = path.join(directory, "source.pptx");
await writeFile(filename, original.bytes);
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, "profile")).href}`;
const nativeLog = await run(executable, [profile, "--headless", "--convert-to", "pptx:Impress Office Open XML", "--outdir", nativeDirectory, filename], options);
const nativeFilename = path.join(nativeDirectory, "source.pptx"), nativeBytes = await readFile(nativeFilename);
const session = await importPptx(nativeBytes, { projection: "content" });
try {
  await writeFile(path.join(directory, "import-report.json"), JSON.stringify(session.report, null, 2) + "\n");
  const noOp = await session.prepareSave({ mode: "no-op" }); assert.equal(noOp.status, "ready");
  assert.equal(Buffer.compare(Buffer.from((await session.save(noOp.plan)).bytes), nativeBytes), 0); noOp.plan.dispose();
  const fresh = await generatePptx(session.document), freshFilename = path.join(directory, "projection.pptx");
  await writeFile(freshFilename, fresh.bytes); assert.deepEqual(fresh.warnings, []);
  const reopenLog = await run(executable, [profile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", reopenedDirectory, freshFilename], options);
  const pdf = await readFile(path.join(reopenedDirectory, "projection.pdf")); assert.ok(pdf.subarray(0, 5).equals(Buffer.from("%PDF-")));
  // Independent stdlib ZIP/XML inventory resolves actual picture relationships;
  // it does not count unrelated media parts as evidence of picture mapping.
  const inventory = JSON.parse((await run("python3", ["-c", `
import sys, zipfile, json, hashlib, posixpath, xml.etree.ElementTree as E
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
result=[]
for filename in sys.argv[1:]:
    images=[]
    with zipfile.ZipFile(filename) as z:
        assert z.testzip() is None
        for part in sorted(n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')):
            relpart=posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
            rels={e.get('Id'): e for e in E.fromstring(z.read(relpart))}
            root=E.fromstring(z.read(part)); parents={child:parent for parent in root.iter() for child in parent}
            for pic in root.findall('.//p:pic',ns):
                name=pic.find('p:nvPicPr/p:cNvPr',ns).get('name')
                blip=pic.find('p:blipFill/a:blip',ns)
                rel=rels[blip.get('{'+r+'}embed')]
                assert rel.get('TargetMode','Internal')=='Internal'
                target=posixpath.normpath(posixpath.join(posixpath.dirname(part),rel.get('Target'))).lstrip('/')
                xfrm=pic.find('p:spPr/a:xfrm',ns); off=xfrm.find('a:off',ns); ext=xfrm.find('a:ext',ns)
                crop=pic.find('p:blipFill/a:srcRect',ns)
                parent=parents[pic]
                origin=parent.find('p:grpSpPr/a:xfrm/a:chOff',ns) if parent.tag=='{'+ns['p']+'}grpSp' else None
                box={k:int(node.get(key))/12700 for k,node,key in [('x',off,'x'),('y',off,'y'),('width',ext,'cx'),('height',ext,'cy')]}
                if origin is not None:
                    box['x']-=int(origin.get('x'))/12700; box['y']-=int(origin.get('y'))/12700
                images.append({'name':name,'sha256':hashlib.sha256(z.read(target)).hexdigest(),
                    'box':box,
                    'crop':{k:int(crop.get(key,'0'))/100000 if crop is not None else 0 for k,key in [('left','l'),('top','t'),('right','r'),('bottom','b')]}})
    result.append(sorted(images,key=lambda e:e['name']))
print(json.dumps(result))
`, nativeFilename, freshFilename], options)).stdout);
  const hash = bytes => createHash("sha256").update(bytes).digest("hex");
  const normalized = normalizePresentation(session.document), coreInventory = [];
  const visit = elements => { for (const e of elements) {
    if (e.type === "group") visit(e.children);
    else if (e.type === "image") coreInventory.push({ name: e.name, sha256: hash(Buffer.from(normalized.assets.find(a => a.id === e.assetId).source.value, "base64")), box: e.box, crop: e.crop });
  } };
  for (const slide of normalized.slides) visit(slide.elements);
  coreInventory.sort((a, b) => a.name.localeCompare(b.name));
  // Source group-origin subtraction and a fresh integer-EMU serialization can
  // differ at floating-point roundoff. Compare at the format's actual unit.
  const atEmuPrecision = rows => rows.map(row => ({ ...row, box: Object.fromEntries(Object.entries(row.box).map(([key, value]) => [key, Math.round(value * 12700)])) }));
  assert.equal(inventory[0].length, 3);
  assert.deepEqual(atEmuPrecision(coreInventory), atEmuPrecision(inventory[0]));
  assert.deepEqual(atEmuPrecision(inventory[1]), atEmuPrecision(inventory[0]));
  const features = {};
  for (const f of session.report.findings) if (f.operation === "import") { const key = `${f.feature}:${f.tier}`; features[key] = (features[key] ?? 0) + 1; }
  const evidence = {
    schemaVersion: 1, recordedAt: new Date().toISOString(), application: version, platform: `${platform()} ${release()} ${arch()}`,
    artifactsDirectory: path.relative(process.cwd(), directory),
    scope: "LibreOffice-produced embedded raster projection: independent image relationship/byte digest/local frame/crop checks, original-byte no-op and fresh projection PDF reopen. Not native visual parity, interactive editing, general image decoding or preservation edit-save.",
    original: { sha256: hash(original.bytes) }, nativeSource: { sha256: hash(nativeBytes), pictures: inventory[0] },
    projection: { assetCount: normalized.assets.length, pictures: coreInventory, importFeatures: features },
    noOpOriginalBytes: true, independentSourceInventoryMapped: true, frameComparison: "Rounded integer EMUs (12700 per point); image digests and crop fractions compare exactly.",
    freshExport: { sha256: hash(fresh.bytes), sourcePictureInventoryRetained: true },
    reopenedPdf: { sha256: hash(pdf), byteLength: pdf.length }, logs: { native: nativeLog, reopen: reopenLog },
  };
  await writeFile("docs/testing/image-import-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
  console.log(JSON.stringify({ directory, pictureCount: coreInventory.length, independentSourceInventoryMapped: true, freshInventoryRetained: true }));
} finally { session.dispose(); }
