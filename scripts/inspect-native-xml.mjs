import { readFile, writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import assert from 'node:assert/strict';
import { inspectZipArchive, inspectXmlPart, inspectPptx, importPptx } from '../packages/pptx-parser/dist/index.js';
import { execFileSync } from 'node:child_process';
const source = JSON.parse(await readFile('docs/testing/libreoffice-evidence.json', 'utf8'));
const fixtures = [];
for (const fixture of source.fixtures) {
  const file = `${source.artifactsDirectory}/roundtrip/${fixture.fixture}.pptx`;
  const bytes = await readFile(file);
  const sha256 = createHash('sha256').update(bytes).digest('hex');
  assert.equal(sha256, fixture.saved.sha256);
  const archive = await inspectZipArchive(bytes);
  const presentation = await inspectPptx(bytes);
  const session = await importPptx(bytes, { projection: "structure" });
  const prepared = await session.prepareSave({ mode: "no-op" });
  assert.equal(prepared.status, "ready");
  const noOp = await session.save(prepared.plan);
  assert.equal(Buffer.compare(Buffer.from(noOp.bytes), bytes), 0);
  const noOpSession = { assurance: noOp.report.assurance, sha256: createHash('sha256').update(noOp.bytes).digest('hex'), sourceSlides: session.document.slides.length, projection: session.report.projection };
  session.dispose();
  const opc = presentation.package;
  assert.equal(presentation.dialect, "transitional");
  assert.equal(presentation.suggestedExtension, "pptx");
  // Extract the already-hashed snapshot, not a second read of a mutable path.
  const decoded = JSON.parse(execFileSync('python3', ['-c', `
import zipfile, json, base64, sys, io, xml.etree.ElementTree as ET
with zipfile.ZipFile(io.BytesIO(sys.stdin.buffer.read())) as archive:
    parts = {name: base64.b64encode(archive.read(name)).decode()
             for name in archive.namelist() if name.endswith((".xml", ".rels"))}
    metadata = {}
    for relation in ET.fromstring(archive.read('_rels/.rels')):
        key = {'core-properties': 'core', 'extended-properties': 'extended'}.get(relation.attrib['Type'].rsplit('/', 1)[-1])
        if key:
            root = ET.fromstring(archive.read(relation.attrib['Target'].lstrip('/')))
            metadata[key] = [[n.tag.rsplit('}', 1)[-1], ''.join(n.itertext())] for n in root if len(n) == 0]
    print(json.dumps({'parts': parts, 'metadata': metadata}))
`], { input: bytes, encoding: 'utf8', timeout: 30000, maxBuffer: 16 * 1024 * 1024 }));
  for (const [key, properties] of Object.entries(decoded.metadata)) {
    for (const [name, value] of properties) {
      const actual = presentation.metadata[key].properties.find((p) => p.name === name);
      assert.ok(actual, `${fixture.fixture}: missing ${key}/${name}`);
      assert.equal(key === "core" ? actual.value : actual.value.value, value, `${fixture.fixture}: ${key}/${name}`);
    }
  }
  const parts = Object.entries(decoded.parts).map(([name, data]) => [name, Buffer.from(data, 'base64')]);
  const reports = [];
  for (const [partName, value] of parts) if (/\.(xml|rels)$/.test(partName)) {
    const inspected = await inspectXmlPart(value, { partName });
    reports.push({ part: partName, root: inspected.root.name, namespaceUri: inspected.root.namespaceUri, characters: inspected.characters, elements: inspected.elements, maxDepth: inspected.maxDepth });
  }
  fixtures.push({ fixture: fixture.fixture, sha256, noOpSession, archiveEntries: archive.entries.length, metadata: presentation.metadata, sectionList: presentation.sectionList, presentationExtensions: presentation.presentationExtensions, slideSize: presentation.slideSize, slides: presentation.slides, structureFindings: presentation.findings, mainPart: presentation.mainPart, dialect: presentation.dialect, documentKind: presentation.kind, macroEnabled: presentation.macroEnabled, opcParts: opc.parts.length, relationships: opc.relationships.length, xmlParts: reports.length, parts: reports });
}
const evidence = { schemaVersion: 1, scope: 'ZIP32, OPC/XML, structure/metadata/sections and explicit structural import with original-byte no-op session output; no drawing import or edit-save claim', producer: source.application, sourceEvidence: 'docs/testing/libreoffice-evidence.json', artifactsDirectory: source.artifactsDirectory, fixtures };
await writeFile('docs/testing/xml-inspection-evidence.json', JSON.stringify(evidence, null, 2) + '\n');
console.log(JSON.stringify({ fixtures: fixtures.length, xmlParts: fixtures.reduce((n,f)=>n+f.xmlParts,0) }));
