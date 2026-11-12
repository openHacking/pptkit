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
import { createNotesImportFixture } from "../tests/fixtures/content-import.mjs";

if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-notes-import.mjs [soffice-executable]");
const run = promisify(execFile), executable = process.argv[2] ?? "soffice", options = { timeout: 45000, maxBuffer: 4 * 1024 * 1024 };
const version = (await run(executable, ["--version"], options)).stdout.trim();
await mkdir(".tmp/notes-import-native", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/notes-import-native/run-"));
const nativeDirectory = path.join(directory, "native"), reopenedDirectory = path.join(directory, "reopened");
await mkdir(nativeDirectory); await mkdir(reopenedDirectory);
const original = await generatePptx(createNotesImportFixture(createPresentation)), filename = path.join(directory, "source.pptx");
await writeFile(filename, original.bytes);
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, "profile")).href}`;
const nativeLog = await run(executable, [profile, "--headless", "--convert-to", "pptx:Impress Office Open XML", "--outdir", nativeDirectory, filename], options);
const nativeFilename = path.join(nativeDirectory, "source.pptx"), nativeBytes = await readFile(nativeFilename);
const session = await importPptx(nativeBytes, { projection: "content" });
try {
  await writeFile(path.join(directory, "import-report.json"), JSON.stringify(session.report, null, 2) + "\n");
  const prepared = await session.prepareSave({ mode: "no-op" }); assert.equal(prepared.status, "ready");
  assert.equal(Buffer.compare(Buffer.from((await session.save(prepared.plan)).bytes), nativeBytes), 0); prepared.plan.dispose();
  const fresh = await generatePptx(session.document), freshFilename = path.join(directory, "projection.pptx");
  assert.deepEqual(fresh.warnings, []); await writeFile(freshFilename, fresh.bytes);
  const reopenLog = await run(executable, [profile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", reopenedDirectory, freshFilename], options);
  const pdf = await readFile(path.join(reopenedDirectory, "projection.pdf")); assert.ok(pdf.subarray(0, 5).equals(Buffer.from("%PDF-")));
  const inventories = JSON.parse((await run("python3", ["-c", `
import sys, zipfile, json, posixpath, xml.etree.ElementTree as E
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
def target(part,value):
    return posixpath.normpath(posixpath.join(posixpath.dirname(part),value)).lstrip('/')
def relationships(z,part):
    name=posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
    return {e.get('Id'):e for e in E.fromstring(z.read(name))} if name in z.namelist() else {}
result=[]
for filename in sys.argv[1:]:
    rows=[]
    with zipfile.ZipFile(filename) as z:
        assert z.testzip() is None
        main='ppt/presentation.xml'; rels=relationships(z,main)
        slides=[target(main,rels[e.get('{'+r+'}id')].get('Target')) for e in E.fromstring(z.read(main)).findall('p:sldIdLst/p:sldId',ns)]
        for slide in slides:
            links=[e for e in relationships(z,slide).values() if e.get('Type').endswith('/notesSlide')]
            assert len(links)<=1
            if not links:
                rows.append({'paragraphs':[''],'actions':[],'bodyCount':0}); continue
            part=target(slide,links[0].get('Target')); rels=relationships(z,part); root=E.fromstring(z.read(part))
            bodies=[sp for sp in root.findall('p:cSld/p:spTree/p:sp',ns) if any(ph.get('type')=='body' for ph in sp.findall('p:nvSpPr/p:nvPr/p:ph',ns))]
            assert len(bodies)<=1
            paragraphs=[]; actions=[]
            if bodies:
                for pi,para in enumerate(bodies[0].findall('p:txBody/a:p',ns)):
                    text=''; runs=[n for n in para if n.tag in ['{'+ns['a']+'}r','{'+ns['a']+'}br']]
                    for ri,node in enumerate(runs):
                        text+='\\n' if node.tag.endswith('}br') else ''.join(t.text or '' for t in node.findall('a:t',ns))
                        click=node.find('a:rPr/a:hlinkClick',ns)
                        if click is not None:
                            rel=rels[click.get('{'+r+'}id')]; kind='url' if rel.get('Type').endswith('/hyperlink') else 'slide'
                            value=rel.get('Target') if kind=='url' else target(part,rel.get('Target'))
                            row={'paragraph':pi,'run':ri,'type':kind,'target':value}
                            if click.get('tooltip') is not None: row['tooltip']=click.get('tooltip')
                            actions.append(row)
                    paragraphs.append(text)
            rows.append({'paragraphs':paragraphs or [''],'actions':actions,'bodyCount':len(bodies)})
    result.append(rows)
print(json.dumps(result))
`, filename, nativeFilename, freshFilename], options)).stdout);
  const normalized = normalizePresentation(session.document);
  const sourceSlides = new Map(normalized.slides.map((s, i) => [s.id, session.report.inspection.slides[i].part.slice(1)]));
  const coreInventory = normalized.slides.map(slide => ({
    paragraphs: slide.notes.map(p => p.runs.map(r => r.text).join("")),
    actions: slide.notes.flatMap((p, pi) => p.runs.flatMap((r, ri) => r.action ? [{ paragraph: pi, run: ri, type: r.action.type, target: r.action.type === "slide" ? sourceSlides.get(r.action.slideId) : r.action.url, ...(r.action.tooltip !== undefined ? { tooltip: r.action.tooltip } : {}) }] : [])),
  }));
  const semantic = rows => rows.map(({ bodyCount, ...value }) => value);
  const nativeUnmappedActions = inventories[1].flatMap((row, slideIndex) => row.actions.filter(a => a.type === "url" && !/^(?:https?:|mailto:)/i.test(a.target)).map(action => ({ slideIndex, ...action })));
  // This producer converts the two notes slide-jumps into relative external
  // hyperlinks. They must remain source-only, not be made actionable by Parser.
  assert.equal(nativeUnmappedActions.length, 2);
  assert.ok(nativeUnmappedActions.every(a => /^\.\.\/slides\/slide[12]\.xml$/.test(a.target)));
  assert.equal(session.report.diagnostics.filter(d => d.code === "notes-click-action-unmapped").length, 2);
  const mappedNative = semantic(inventories[1]).map(row => ({ ...row, actions: row.actions.filter(a => a.type === "slide" || /^(?:https?:|mailto:)/i.test(a.target)) }));
  assert.equal(inventories[1].length, 3);
  assert.ok(inventories[1][0].paragraphs.join("\n").includes("Speaker α😀 & <literal>"));
  assert.deepEqual(coreInventory, mappedNative); assert.deepEqual(semantic(inventories[2]), coreInventory);
  const hash = bytes => createHash("sha256").update(bytes).digest("hex"), findings = {};
  for (const f of session.report.findings) if (f.operation === "import" && f.feature.startsWith("notes.")) { const key = `${f.feature}:${f.tier}`; findings[key] = (findings[key] ?? 0) + 1; }
  const evidence = {
    schemaVersion: 1, recordedAt: new Date().toISOString(), application: version, platform: `${platform()} ${release()} ${arch()}`,
    artifactsDirectory: path.relative(process.cwd(), directory),
    scope: "LibreOffice-produced direct notes-body text/click projection: independent notes XML text and click inventory, exact no-op bytes, fresh projection export and slide PDF reopen. Not notes-page geometry, master inheritance, native notes rendering, interactive editing/navigation or preservation edit-save.",
    original: { sha256: hash(original.bytes), inventory: inventories[0] }, nativeSource: { sha256: hash(nativeBytes), inventory: inventories[1] },
    projection: { inventory: coreInventory, importFindings: findings, sourceOnlyActions: nativeUnmappedActions }, independentMappedInventoryVerified: true, noOpOriginalBytes: true,
    producerChangedInventory: JSON.stringify(inventories[0]) !== JSON.stringify(inventories[1]),
    producerLimitations: "LibreOffice splits the authored line break into paragraphs, converts notes slide-jumps into relative external links, strips tooltips and normalizes URL escaping. The two relative links remain source-only with located diagnostics; fresh export retains projected text and supported actions only.",
    freshExport: { sha256: hash(fresh.bytes), inventory: inventories[2], sourceTextInventoryRetained: true, mappedActionInventoryRetained: true, sourceOnlyActionCount: nativeUnmappedActions.length },
    reopenedPdf: { sha256: hash(pdf), byteLength: pdf.length }, logs: { native: nativeLog, reopen: reopenLog },
  };
  await writeFile("docs/testing/notes-import-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
  console.log(JSON.stringify({ directory, slides: coreInventory.length, notesBodies: inventories[1].filter(r => r.bodyCount).length, paragraphs: coreInventory.map(r => r.paragraphs.length), actions: coreInventory.map(r => r.actions.length), sourceOnlyActions: nativeUnmappedActions.length, independentMappedInventoryVerified: true }));
} finally { session.dispose(); }
