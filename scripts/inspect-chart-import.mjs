// Optional native producer experiment; interactive Edit Data remains untested.
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdtemp, mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { createHash } from "node:crypto";
import { platform, release, arch } from "node:os";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { createChartImportFixture } from "../tests/fixtures/content-import.mjs";

if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-chart-import.mjs [soffice-executable]");
const run = promisify(execFile), executable = process.argv[2] ?? "soffice", options = { timeout: 45000, maxBuffer: 8 * 1024 * 1024 };
const version = (await run(executable, ["--version"], options)).stdout.trim();
await mkdir(".tmp/chart-import-native", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/chart-import-native/run-"));
const nativeDirectory = path.join(directory, "native"), reopenedDirectory = path.join(directory, "reopened");
await mkdir(nativeDirectory); await mkdir(reopenedDirectory);
const original = await generatePptx(createChartImportFixture(createPresentation)), filename = path.join(directory, "source.pptx");
assert.deepEqual(original.warnings, []); await writeFile(filename, original.bytes);
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, "profile")).href}`;
const nativeLog = await run(executable, [profile, "--headless", "--convert-to", "pptx:Impress Office Open XML", "--outdir", nativeDirectory, filename], options);
const nativeFilename = path.join(nativeDirectory, "source.pptx"), nativeBytes = await readFile(nativeFilename);
const session = await importPptx(nativeBytes, { projection: "content" });
try {
  await writeFile(path.join(directory, "import-report.json"), JSON.stringify(session.report, null, 2) + "\n");
  const prepared = await session.prepareSave({ mode: "no-op" }); assert.equal(prepared.status, "ready");
  assert.equal(Buffer.compare(Buffer.from((await session.save(prepared.plan)).bytes), nativeBytes), 0);
  const fresh = await generatePptx(session.document), freshFilename = path.join(directory, "projection.pptx");
  assert.deepEqual(fresh.warnings, []); await writeFile(freshFilename, fresh.bytes);
  const reopenLog = await run(executable, [profile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", reopenedDirectory, freshFilename], options);
  const pdf = await readFile(path.join(reopenedDirectory, "projection.pdf")); assert.ok(pdf.subarray(0, 5).equals(Buffer.from("%PDF-")));
  const inventories = JSON.parse((await run("python3", ["-c", `
import sys, zipfile, json, posixpath, xml.etree.ElementTree as E
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
def target(part,value): return posixpath.normpath(posixpath.join(posixpath.dirname(part),value)).lstrip('/')
def relationships(z,part):
    name=posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
    return {e.get('Id'):e for e in E.fromstring(z.read(name))} if name in z.namelist() else {}
def values(node,numeric=False):
    if node is None: return []
    direct=node.find('c:v',ns)
    if direct is not None: return [direct.text or '']
    points=node.findall('.//c:pt',ns)
    return [(float(p.find('c:v',ns).text) if numeric else p.find('c:v',ns).text or '') for p in sorted(points,key=lambda p:int(p.get('idx')))]
result=[]
for filename in sys.argv[1:]:
    output=[]
    with zipfile.ZipFile(filename) as z:
        assert z.testzip() is None
        main='ppt/presentation.xml'; rels=relationships(z,main)
        slides=[target(main,rels[e.get('{'+r+'}id')].get('Target')) for e in E.fromstring(z.read(main)).findall('p:sldIdLst/p:sldId',ns)]
        for slide in slides:
            slide_rels=relationships(z,slide)
            for ref in E.fromstring(z.read(slide)).findall('.//c:chart',ns):
                part=target(slide,slide_rels[ref.get('{'+r+'}id')].get('Target'))
                root=E.fromstring(z.read(part)); plot=root.find('c:chart/c:plotArea',ns)
                families=[n for n in plot if n.tag.endswith('Chart')]
                assert len(families)==1
                family=families[0]; data=[]
                for ser in sorted(family.findall('c:ser',ns),key=lambda s:int(s.find('c:order',ns).get('val'))):
                    data.append({'name':values(ser.find('c:tx',ns))[0],'categories':values(ser.find('c:cat',ns)),'values':values(ser.find('c:val',ns),True)})
                output.append({'chartType':family.tag.split('}')[1][:-5],'series':data})
    result.append(output)
print(json.dumps(result))
`, filename, nativeFilename, freshFilename], options)).stdout);
  const core = session.document.slides.flatMap(s => s.elements.filter(e => e.type === "chart").map(e => ({ chartType: e.chartType, series: e.series.map(s => ({ name: s.name, categories: e.categories, values: s.values })) })));
  assert.equal(core.length, 3, "all three native chart families remain independently identifiable");
  assert.deepEqual(core, inventories[1], "Core equals native cache inventory from independent ZIP/XML reader");
  assert.deepEqual(inventories[2], core, "fresh native cache data equals projection");
  const hash = bytes => createHash("sha256").update(bytes).digest("hex");
  const evidence = {
    schemaVersion: 1, recordedAt: new Date().toISOString(), application: version, platform: `${platform()} ${release()} ${arch()}`,
    artifactsDirectory: path.relative(process.cwd(), directory),
    scope: "LibreOffice-produced bar/line/pie data import, independent Python ZIP/XML cache inventory, byte-exact no-op, fresh export and PDF reopen. Does not verify interactive Edit Data, PowerPoint, visual fidelity, formula evaluation or preservation edits.",
    original: { sha256: hash(original.bytes), inventory: inventories[0] }, nativeSource: { sha256: hash(nativeBytes), inventory: inventories[1] },
    projection: { inventory: core }, independentMappedInventoryVerified: true, noOpOriginalBytes: true,
    producerChangedData: JSON.stringify(inventories[0]) !== JSON.stringify(inventories[1]),
    freshExport: { sha256: hash(fresh.bytes), inventory: inventories[2] }, reopenedPdf: { sha256: hash(pdf), byteLength: pdf.length },
    logs: { native: nativeLog, reopen: reopenLog },
  };
  await writeFile("docs/testing/chart-import-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
  console.log(JSON.stringify({ directory, charts: core.length, producerChangedData: evidence.producerChangedData, independentMappedInventoryVerified: true }));
} finally { session.dispose(); }
