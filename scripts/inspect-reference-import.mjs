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
import { createReferenceImportFixture } from "../tests/fixtures/content-import.mjs";

if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-reference-import.mjs [soffice-executable]");
const run = promisify(execFile), executable = process.argv[2] ?? "soffice", options = { timeout: 45000, maxBuffer: 4 * 1024 * 1024 };
const version = (await run(executable, ["--version"], options)).stdout.trim();
await mkdir(".tmp/reference-import-native", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/reference-import-native/run-"));
const nativeDirectory = path.join(directory, "native"), reopenedDirectory = path.join(directory, "reopened");
await mkdir(nativeDirectory); await mkdir(reopenedDirectory);
const original = await generatePptx(createReferenceImportFixture(createPresentation), { connectorMode: "native" }), filename = path.join(directory, "source.pptx");
await writeFile(filename, original.bytes);
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, "profile")).href}`;
const nativeLog = await run(executable, [profile, "--headless", "--convert-to", "pptx:Impress Office Open XML", "--outdir", nativeDirectory, filename], options);
const nativeFilename = path.join(nativeDirectory, "source.pptx"), nativeBytes = await readFile(nativeFilename);
const session = await importPptx(nativeBytes, { projection: "content" });
try {
  await writeFile(path.join(directory, "import-report.json"), JSON.stringify(session.report, null, 2) + "\n");
  const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready");
  assert.equal(Buffer.compare(Buffer.from((await session.save(plan.plan)).bytes), nativeBytes), 0); plan.plan.dispose();
  const fresh = await generatePptx(session.document, { connectorMode: "native" }), freshFilename = path.join(directory, "projection.pptx");
  assert.deepEqual(fresh.warnings, []); await writeFile(freshFilename, fresh.bytes);
  const reopenLog = await run(executable, [profile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", reopenedDirectory, freshFilename], options);
  const pdf = await readFile(path.join(reopenedDirectory, "projection.pdf")); assert.ok(pdf.subarray(0, 5).equals(Buffer.from("%PDF-")));
  const inventory = JSON.parse((await run("python3", ["-c", `
import sys, zipfile, json, posixpath, xml.etree.ElementTree as E
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
result=[]
for filename in sys.argv[1:]:
    connections=[]; actions=[]
    with zipfile.ZipFile(filename) as z:
        assert z.testzip() is None
        for part in sorted(n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')):
            root=E.fromstring(z.read(part)); names={e.get('id'):e.get('name') for e in root.findall('.//p:cNvPr',ns)}
            relpart=posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
            rels={e.get('Id'):e for e in E.fromstring(z.read(relpart))}
            def action(owner, location):
                click=owner.find('a:hlinkClick',ns)
                if click is None: return
                rel=rels[click.get('{'+r+'}id')]
                target=rel.get('Target')
                kind='url' if rel.get('Type').endswith('/hyperlink') else 'slide'
                if kind=='slide': target=posixpath.normpath(posixpath.join(posixpath.dirname(part),target)).lstrip('/')
                row={'location':location,'type':kind,'target':target}
                if click.get('tooltip') is not None: row['tooltip']=click.get('tooltip')
                actions.append(row)
            for drawing in root.iter():
                nv=next((child for child in drawing if child.tag in ['{'+ns['p']+'}'+n for n in ['nvSpPr','nvCxnSpPr','nvPicPr','nvGrpSpPr']]),None)
                if nv is None: continue
                cnv=nv.find('p:cNvPr',ns); name=cnv.get('name')
                action(cnv,name+'/object')
                body=drawing.find('p:txBody',ns)
                if body is not None:
                    for pi,para in enumerate(body.findall('a:p',ns)):
                        runs=[n for n in para if n.tag in ['{'+ns['a']+'}r','{'+ns['a']+'}br']]
                        for ri,node in enumerate(runs):
                            props=node.find('a:rPr',ns)
                            if props is not None: action(props,name+'/paragraph-'+str(pi)+'/run-'+str(ri))
                cxn=nv.find('p:cNvCxnSpPr',ns)
                if cxn is not None:
                    for side,tag in [('start','stCxn'),('end','endCxn')]:
                        endpoint=cxn.find('a:'+tag,ns)
                        if endpoint is not None: connections.append({'name':name,'side':side,'target':names[endpoint.get('id')],'site':int(endpoint.get('idx'))})
    result.append({'connections':sorted(connections,key=lambda e:(e['name'],e['side'])),'actions':sorted(actions,key=lambda e:e['location'])})
print(json.dumps(result))
`, filename, nativeFilename, freshFilename], options)).stdout);
  const normalized = normalizePresentation(session.document), byId = new Map(), slideParts = new Map(normalized.slides.map((s, i) => [s.id, session.report.inspection.slides[i].part.slice(1)]));
  const collect = elements => { for (const e of elements) { byId.set(e.id, e); if (e.type === "group") collect(e.children); } };
  normalized.slides.forEach(s => collect(s.elements));
  const connections = [], actions = [];
  const action = (a, location) => { if (a) actions.push({ location, type: a.type, target: a.type === "slide" ? slideParts.get(a.slideId) : a.url, ...(a.tooltip !== undefined ? { tooltip: a.tooltip } : {}) }); };
  for (const e of byId.values()) {
    action(e.action, `${e.name}/object`);
    const text = e.type === "text" ? e.content : e.type === "shape" ? e.text?.content : undefined;
    text?.forEach((p, pi) => p.runs.forEach((r, ri) => action(r.action, `${e.name}/paragraph-${pi}/run-${ri}`)));
    if (e.type === "connector") for (const side of ["start", "end"]) if ("elementId" in e[side]) {
      const target = byId.get(e[side].elementId), site = ["top", "left", "bottom", "right"].indexOf(e[side].anchor) * (target.type === "shape" && target.shape === "ellipse" ? 2 : 1);
      connections.push({ name: e.name, side, target: target.name, site });
    }
  }
  const coreInventory = { connections: connections.sort((a, b) => a.name.localeCompare(b.name) || a.side.localeCompare(b.side)), actions: actions.sort((a, b) => a.location.localeCompare(b.location)) };
  assert.equal(inventory[1].connections.length, 3);
  assert.equal(inventory[1].actions.length, 4, JSON.stringify(inventory));
  assert.deepEqual(coreInventory, inventory[1]); assert.deepEqual(inventory[2], inventory[1]);
  const hash = bytes => createHash("sha256").update(bytes).digest("hex"), features = {};
  for (const f of session.report.findings) if (f.operation === "import") { const key = `${f.feature}:${f.tier}`; features[key] = (features[key] ?? 0) + 1; }
  const evidence = {
    schemaVersion: 1, recordedAt: new Date().toISOString(), application: version, platform: `${platform()} ${release()} ${arch()}`,
    artifactsDirectory: path.relative(process.cwd(), directory),
    scope: "LibreOffice-produced native straight connector/explicit click projection: independent drawing-name/site and hyperlink/slide/tooltip inventories, exact no-op bytes and fresh native projection PDF reopen. Not interactive attachment edits, actual click navigation, native visual parity or preservation edit-save.",
    original: { sha256: hash(original.bytes), inventory: inventory[0] }, nativeSource: { sha256: hash(nativeBytes), inventory: inventory[1] },
    projection: { inventory: coreInventory, importFeatures: features },
    producerDifferences: {
      omittedOriginalActionLocations: inventory[0].actions.filter((sourceAction) => !inventory[1].actions.some((nativeAction) => nativeAction.location === sourceAction.location)).map(({ location }) => location),
      changedOriginalActions: inventory[0].actions.flatMap((sourceAction) => {
        const nativeAction = inventory[1].actions.find(({ location }) => location === sourceAction.location);
        return nativeAction && JSON.stringify(nativeAction) !== JSON.stringify(sourceAction) ? [{ location: sourceAction.location, original: sourceAction, native: nativeAction }] : [];
      }),
      retainedNativeActions: inventory[1].actions.length,
      note: "The native conversion dropped connector/group object actions and removed tooltips from surviving actions; URL escaping was normalized. This evidence does not attribute those producer changes to PPTKit.",
    },
    noOpOriginalBytes: true, independentSourceInventoryMapped: true,
    freshExport: { sha256: hash(fresh.bytes), sourceReferenceInventoryRetained: true },
    reopenedPdf: { sha256: hash(pdf), byteLength: pdf.length }, logs: { native: nativeLog, reopen: reopenLog },
  };
  await writeFile("docs/testing/reference-import-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
  console.log(JSON.stringify({ directory, connections: connections.length, actions: actions.length, independentSourceInventoryMapped: true, freshInventoryRetained: true }));
} finally { session.dispose(); }
