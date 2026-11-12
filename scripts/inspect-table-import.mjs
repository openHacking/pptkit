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
import { createTableImportFixture } from "../tests/fixtures/content-import.mjs";

if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-table-import.mjs [soffice-executable]");
const run = promisify(execFile), executable = process.argv[2] ?? "soffice";
const options = { timeout: 45000, maxBuffer: 8 * 1024 * 1024 };
const version = (await run(executable, ["--version"], options)).stdout.trim();
await mkdir(".tmp/table-import-native", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/table-import-native/run-"));
const nativeDirectory = path.join(directory, "native"), reopenedDirectory = path.join(directory, "reopened");
await mkdir(nativeDirectory); await mkdir(reopenedDirectory);
const original = await generatePptx(createTableImportFixture(createPresentation)), filename = path.join(directory, "source.pptx");
assert.deepEqual(original.warnings, []); await writeFile(filename, original.bytes);
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
  const pdf = await readFile(path.join(reopenedDirectory, "projection.pdf"));
  assert.ok(pdf.subarray(0, 5).equals(Buffer.from("%PDF-")));
  // stdlib ZIP/XML inventory is independent of PPTKit's parser and writer.
  const inventories = JSON.parse((await run("python3", ["-c", `
import sys, zipfile, json, posixpath, xml.etree.ElementTree as E
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
def target(part,value): return posixpath.normpath(posixpath.join(posixpath.dirname(part),value)).lstrip('/')
def relationships(z,part):
    name=posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
    return {e.get('Id'):e for e in E.fromstring(z.read(name))} if name in z.namelist() else {}
def boolean(value): return value in ('true','1')
result=[]
for filename in sys.argv[1:]:
    output=[]
    with zipfile.ZipFile(filename) as z:
        assert z.testzip() is None
        main='ppt/presentation.xml'; rels=relationships(z,main)
        slides=[target(main,rels[e.get('{'+r+'}id')].get('Target')) for e in E.fromstring(z.read(main)).findall('p:sldIdLst/p:sldId',ns)]
        for slide_index,slide in enumerate(slides):
            root=E.fromstring(z.read(slide)); parents={c:p for p in root.iter() for c in p}
            for frame in root.findall('.//p:graphicFrame',ns):
                table=frame.find('a:graphic/a:graphicData/a:tbl',ns)
                if table is None: continue
                xfrm=frame.find('p:xfrm',ns); off=xfrm.find('a:off',ns); ext=xfrm.find('a:ext',ns)
                origin={'x':0,'y':0}; parent=parents.get(frame)
                if parent is not None and parent.tag=='{'+ns['p']+'}grpSp':
                    child_origin=parent.find('p:grpSpPr/a:xfrm/a:chOff',ns)
                    if child_origin is not None: origin={k:int(child_origin.get(k)) for k in ('x','y')}
                rows=table.findall('a:tr',ns); anchors=[]; physical=[]
                for ri,row in enumerate(rows):
                    physical.append([])
                    for ci,cell in enumerate(row.findall('a:tc',ns)):
                        physical[-1].append({k:cell.get(k) for k in ('rowSpan','gridSpan','hMerge','vMerge') if cell.get(k) is not None})
                        if boolean(cell.get('hMerge')) or boolean(cell.get('vMerge')): continue
                        paragraphs=[]
                        for para in cell.findall('a:txBody/a:p',ns):
                            text=''
                            for child in para:
                                if child.tag=='{'+ns['a']+'}br': text+='\\n'
                                elif child.tag=='{'+ns['a']+'}r': text+=''.join(t.text or '' for t in child.findall('a:t',ns))
                            paragraphs.append(text)
                        anchors.append({'row':ri,'column':ci,'rowSpan':int(cell.get('rowSpan','1')),'colSpan':int(cell.get('gridSpan','1')),'paragraphs':paragraphs or ['']})
                value={'slide':slide_index,'box':{'x':(int(off.get('x'))-origin['x'])/12700,'y':(int(off.get('y'))-origin['y'])/12700,'width':int(ext.get('cx'))/12700,'height':int(ext.get('cy'))/12700},
                    'transform':{'rotation':int(xfrm.get('rot','0'))/60000,'flipH':boolean(xfrm.get('flipH')),'flipV':boolean(xfrm.get('flipV'))},
                    'columns':[int(c.get('w'))/12700 for c in table.findall('a:tblGrid/a:gridCol',ns)],'heights':[int(row.get('h'))/12700 for row in rows],'anchors':anchors,'physical':physical}
                output.append(value)
    result.append(output)
print(json.dumps(result))
`, filename, nativeFilename, freshFilename], options)).stdout);
  const core = [], normalized = normalizePresentation(session.document);
  const visit = (element, slideIndex) => {
    if (element.type === "group") { element.children.forEach(child => visit(child, slideIndex)); return; }
    if (element.type !== "table") return;
    // Build a brute-force anchor grid from compact Core rows, independently of
    // the parser's one-row active-merge algorithm.
    const grid = element.rows.map(() => element.columns.map(() => false)), anchors = [];
    element.rows.forEach((row, ri) => {
      let column = 0;
      for (const cell of row.cells) {
        while (grid[ri][column]) column++;
        anchors.push({ row: ri, column, rowSpan: cell.rowSpan, colSpan: cell.colSpan, paragraphs: cell.content.map(p => p.runs.map(r => r.text).join("")) });
        for (let y = ri; y < ri + cell.rowSpan; y++) for (let x = column; x < column + cell.colSpan; x++) {
          assert.equal(grid[y][x], false); grid[y][x] = true;
        }
        column += cell.colSpan;
      }
    });
    core.push({ slide: slideIndex, box: element.box, transform: element.transform, columns: element.columns, heights: element.rows.map(r => r.height), anchors });
  };
  normalized.slides.forEach((slide, i) => slide.elements.forEach(element => visit(element, i)));
  const semantic = values => values.map(({ physical, ...value }) => value);
  const integerGeometry = values => values.map(value => ({
    ...value,
    box: Object.fromEntries(Object.entries(value.box).map(([key, n]) => [key, Math.round(n * 12700)])),
    columns: value.columns.map(n => Math.round(n * 12700)),
    heights: value.heights.map(n => Math.round(n * 12700)),
    transform: { ...value.transform, rotation: Math.round(value.transform.rotation * 60000) },
  }));
  assert.equal(core.length, 2, "both fixture tables must remain independently identifiable");
  assert.deepEqual(integerGeometry(core), integerGeometry(semantic(inventories[1])), "Core matches native grid/anchors/text and exact integer-EMU geometry");
  assert.deepEqual(integerGeometry(semantic(inventories[2])), integerGeometry(core), "fresh output matches projected semantics");
  const hash = bytes => createHash("sha256").update(bytes).digest("hex"), importFindings = {};
  for (const f of session.report.findings) if (f.operation === "import" && (f.feature === "drawing.table.mapped-fields" || /a:tbl/.test(f.source.path ?? ""))) {
    const key = `${f.feature}:${f.tier}`; importFindings[key] = (importFindings[key] ?? 0) + 1;
  }
  const evidence = {
    schemaVersion: 1, recordedAt: new Date().toISOString(), application: version, platform: `${platform()} ${release()} ${arch()}`,
    artifactsDirectory: path.relative(process.cwd(), directory),
    scope: "LibreOffice-produced table grid/merge/anchor-text/local-transform projection, independent Python ZIP/XML inventory, exact no-op bytes, fresh projection export and PDF reopen. Not native table style fidelity, frame/grid mismatch layout, interactive table editing, PowerPoint verification or preservation edit-save.",
    original: { sha256: hash(original.bytes), inventory: inventories[0] },
    nativeSource: { sha256: hash(nativeBytes), inventory: inventories[1] },
    projection: { inventory: core, importFindings },
    independentMappedInventoryVerified: true, noOpOriginalBytes: true,
    geometryComparison: "Integer EMU positions/sizes and 1/60000-degree rotation; point subtraction order may differ below 1 EMU.",
    producerChangedInventory: JSON.stringify(integerGeometry(semantic(inventories[0]))) !== JSON.stringify(integerGeometry(semantic(inventories[1]))),
    producerGridSummary: inventories.slice(0, 2).map(tables => tables.map(t => ({ columns: t.columns.length, rows: t.heights.length, anchors: t.anchors.length, box: t.box, transform: t.transform }))),
    freshExport: { sha256: hash(fresh.bytes), inventory: inventories[2], mappedGridMergeTextTransformRetained: true },
    reopenedPdf: { sha256: hash(pdf), byteLength: pdf.length }, logs: { native: nativeLog, reopen: reopenLog },
  };
  await writeFile("docs/testing/table-import-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
  console.log(JSON.stringify({ directory, tables: core.length, anchors: core.map(t => t.anchors.length), producerChangedInventory: evidence.producerChangedInventory, independentMappedInventoryVerified: true }));
} finally { session.dispose(); }
