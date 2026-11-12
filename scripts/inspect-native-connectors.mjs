// Optional native application evidence; not required by portable CI.
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
import { createNativeConnectorFixture } from "../tests/fixtures/native-connectors.mjs";
if (process.argv.length > 3) throw new Error("Usage: node scripts/inspect-native-connectors.mjs [soffice-executable]");
const run = promisify(execFile), executable = process.argv[2] ?? "soffice", options = { timeout: 45000, maxBuffer: 4 * 1024 * 1024 };
const version = (await run(executable, ["--version"], options)).stdout.trim();
await mkdir(".tmp/native-connectors", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/native-connectors/run-")), output = path.join(directory, "saved"), reopened = path.join(directory, "reopened");
await mkdir(output); await mkdir(reopened);
const input = path.join(directory, "connectors.pptx"), generated = await generatePptx(createNativeConnectorFixture(createPresentation), { connectorMode: "native" });
assert.deepEqual(generated.warnings, []); await writeFile(input, generated.bytes);
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, "profile")).href}`;
const saveLog = await run(executable, [profile, "--headless", "--convert-to", "pptx:Impress Office Open XML", "--outdir", output, input], options);
const savedPath = path.join(output, "connectors.pptx"), savedBytes = await readFile(savedPath);
const reopenLog = await run(executable, [profile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", reopened, savedPath], options);
const pdf = await readFile(path.join(reopened, "connectors.pdf")); assert.ok(pdf.subarray(0, 5).equals(Buffer.from("%PDF-")));
const independent = JSON.parse((await run("python3", ["-c", `
import sys, zipfile, json, xml.etree.ElementTree as E
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
result=[]
for filename in sys.argv[1:]:
    connectors=[]
    with zipfile.ZipFile(filename) as z:
        assert z.testzip() is None
        for part in sorted(n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')):
            root=E.fromstring(z.read(part))
            names={n.get('id'):n.get('name') for n in root.findall('.//p:cNvPr',ns)}
            for c in root.findall('.//p:cxnSp',ns):
                props=c.find('p:nvCxnSpPr/p:cNvPr',ns)
                record={'part':part,'name':props.get('name'),'id':props.get('id'),'bindings':{}}
                for side in ['stCxn','endCxn']:
                    node=c.find('p:nvCxnSpPr/p:cNvCxnSpPr/a:'+side,ns)
                    if node is not None:
                        assert node.get('id') in names
                        record['bindings'][side]={'target':names[node.get('id')],'index':int(node.get('idx'))}
                preset=c.find('p:spPr/a:prstGeom',ns)
                record['preset']=None if preset is None else preset.get('prst')
                xfrm=c.find('p:spPr/a:xfrm',ns)
                record['transform']=None if xfrm is None else {'attributes':dict(xfrm.attrib),'children':[{'tag':n.tag.split('}')[-1],**n.attrib} for n in xfrm]}
                connectors.append(record)
    result.append(connectors)
print(json.dumps(result))
`, input, savedPath], options)).stdout);
const [before, after] = independent;
assert.equal(before.length, 10); assert.equal(before.filter(c => Object.keys(c.bindings).length).length, 5);
const summary = items => items.map(({ part, name, bindings }) => ({ part, name, bindings })).sort((a, b) => a.name.localeCompare(b.name));
const preserved = JSON.stringify(summary(before)) === JSON.stringify(summary(after));
const hash = bytes => createHash("sha256").update(bytes).digest("hex");
const evidence = { schemaVersion: 1, application: version, platform: `${platform()} ${release()} ${arch()}`, artifactsDirectory: path.relative(process.cwd(), directory), scope: "Experimental native straight connector profile; headless save/reopen and independent attachment inventory, not interactive attachment movement, visual parity or PowerPoint certification", original: { sha256: hash(generated.bytes), connectors: before }, saved: { sha256: hash(savedBytes), connectors: after }, reopenedPdf: { sha256: hash(pdf), byteLength: pdf.length }, attachmentInventoryPreserved: preserved, logs: { save: saveLog, reopen: reopenLog } };
await writeFile("docs/testing/native-connectors-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ directory, original: before.length, saved: after.length, attachmentInventoryPreserved: preserved }));
