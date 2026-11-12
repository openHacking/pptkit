// Optional native experiment; independent Python ZIP/XML is the timing oracle.
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { platform, release, arch } from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx, inspectPptx } from "../packages/pptx-parser/dist/index.js";

const args = process.argv.slice(2), all = args.includes("--all"), positional = args.filter(a => a !== "--all");
if (positional.length > 1 || args.filter(a => a === "--all").length > 1) throw new Error("Usage: node scripts/inspect-slide-transitions.mjs [--all] [soffice-executable]");
const executable = positional[0] ?? "soffice", run = promisify(execFile), options = { timeout: 45000, maxBuffer: 8 * 1024 * 1024 };
const application = (await run(executable, ["--version"], options)).stdout.trim();
await mkdir(".tmp/transition-native", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/transition-native/run-"));
const native = path.join(directory, "native"), pdf = path.join(directory, "pdf");
await mkdir(native); await mkdir(pdf);
const transitions = [
  { type: "fade", durationMs: 1234, onClick: true, afterMs: 3456 },
  { type: "fade", durationMs: 0, onClick: false, afterMs: 0 },
  { type: "fade", durationMs: 2000, onClick: false },
];
if (all) for (const type of ["push", "wipe"]) for (const direction of ["left", "right", "up", "down"]) transitions.push({ type, direction, durationMs: 765, onClick: false, afterMs: 2345 });
const deck = createPresentation();
for (const [index, transition] of transitions.entries()) deck.addSlide({ transition }).addElement({ type: "text", content: `${transition.type} ${transition.direction ?? "plain"} control ${index + 1}`, box: { x: 40, y: 40, width: 400, height: 60 } });
const bytes = (await generatePptx(deck)).bytes, source = path.join(directory, "source.pptx");
await writeFile(source, bytes);
const session = await importPptx(bytes, { projection: "content" });
try {
  const p = await session.prepareSave({ mode: "no-op" }); assert.equal(p.status, "ready");
  assert.deepEqual((await session.save(p.plan)).bytes, bytes);
} finally { session.dispose(); }
const profile = `-env:UserInstallation=${pathToFileURL(path.join(directory, "profile")).href}`;
const logs = {
  save: await run(executable, [profile, "--headless", "--convert-to", "pptx:Impress Office Open XML", "--outdir", native, source], options),
  reopen: await run(executable, [profile, "--headless", "--convert-to", "pdf:impress_pdf_Export", "--outdir", pdf, path.join(native, "source.pptx")], options),
};
const python = String.raw`
import sys,json,zipfile,xml.etree.ElementTree as E,fitz
expected=json.loads(sys.argv[3])
P='http://schemas.openxmlformats.org/presentationml/2006/main'
D='http://schemas.microsoft.com/office/powerpoint/2010/main'
def inspect(file):
  with zipfile.ZipFile(file) as z:
    result=[]
    for i in range(1,len(expected)+1):
      root=E.fromstring(z.read('ppt/slides/slide%d.xml'%i))
      result.append([{'attributes':dict(n.attrib),'durationMs':n.get('{'+D+'}dur'),'onClick':n.get('advClick','1'),'afterMs':n.get('advTm'),'effects':[c.tag for c in n],'directions':[c.get('dir') for c in n]} for n in root.iter('{'+P+'}transition')])
    return result
a,b=inspect(sys.argv[1]),inspect(sys.argv[2])
for item,wanted in zip(a,expected):
  assert len(item)==1
  n=item[0]
  assert n['durationMs']==str(wanted['durationMs']) and n['onClick']==('1' if wanted['onClick'] else '0') and n['afterMs']==(str(wanted['afterMs']) if 'afterMs' in wanted else None)
  assert n['effects']==['{'+P+'}'+wanted['type']]
  assert n['directions']==[{'left':'l','right':'r','up':'u','down':'d'}.get(wanted.get('direction'))]
pdf=fitz.open(sys.argv[4])
assert len(pdf)==len(expected)
text='\n'.join(page.get_text() for page in pdf)
for i in range(1,len(expected)+1): assert ('control %d'%i) in text
print(json.dumps({'oracle':'Python zipfile + ElementTree','source':a,'native':b,'rawTransitionRecordsEqual':a==b,'reopenedPdfPages':len(pdf)}))
`;
const independent = JSON.parse((await run("python3", ["-c", python, source, path.join(native, "source.pptx"), JSON.stringify(transitions), path.join(pdf, "source.pdf")], options)).stdout);
const nativeBytes = await readFile(path.join(native, "source.pptx")), pdfBytes = await readFile(path.join(pdf, "source.pdf"));
assert.ok(pdfBytes.length > 0);
const hash = value => createHash("sha256").update(value).digest("hex");
const timingComparisons = transitions.map((expected, i) => {
  const nativeRecords = independent.native[i], exact = nativeRecords.find(n => n.durationMs !== null);
  const direction = { left: "l", right: "r", up: "u", down: "d" }[expected.direction] ?? null;
  return { slide: i + 1, type: expected.type, ...(expected.direction ? { direction: expected.direction } : {}),
    effectAndDirectionRetained: nativeRecords.length > 0 && nativeRecords.every(n => JSON.stringify(n.effects) === JSON.stringify(["{http://schemas.openxmlformats.org/presentationml/2006/main}" + expected.type]) && JSON.stringify(n.directions) === JSON.stringify([direction])),
    nativeDurationMs: exact?.durationMs ?? null,
    exactDurationRetained: exact?.durationMs === String(expected.durationMs),
    clickPolicyRetained: nativeRecords.length > 0 && nativeRecords.every(n => n.onClick === (expected.onClick ? "1" : "0")),
    automaticAdvanceRetained: nativeRecords.length > 0 && nativeRecords.every(n => n.afterMs === (expected.afterMs === undefined ? null : String(expected.afterMs))) };
});
const evidence = {
  schemaVersion: 1, scope: `${transitions.length} named transition records, headless native save and reopening of saved output to PDF. Not interactive settings/editability, actual slideshow playback, repair-dialog absence, PowerPoint or visual fidelity.`,
  environment: { platform: platform(), release: release(), arch: arch(), application }, directory,
  sourceSha256: hash(bytes), nativeSha256: hash(nativeBytes), reopenedPdfSha256: hash(pdfBytes), transitions, independent, timingComparisons,
  parserNativeTransitions: (await inspectPptx(nativeBytes)).slides.map(s => s.transition ?? null), logs,
};
await writeFile(all ? "docs/testing/side-transition-evidence.json" : "docs/testing/slide-transition-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ directory, application, timingComparisons, parserNativeTransitions: evidence.parserNativeTransitions }, null, 2));
