// Optional adversarial checks for retained local native evidence; no UI operations.
import assert from "node:assert/strict";
import { mkdtemp, cp, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { createHash } from "node:crypto";
const run = promisify(execFile), directory = path.resolve(process.argv[2] ?? ".tmp/smartart-stale-cache/powerpoint-2zdvkchy");
const plan = JSON.parse(await readFile(path.join(directory, "plan.json"))), sha = b => createHash("sha256").update(b).digest("hex");
const original = await readFile(path.join(directory, "process-baseline-saved.pptx")), evidence = await readFile(new URL("../docs/testing/smartart-stale-cache-powerpoint-evidence.json", import.meta.url));
const temporary = await mkdtemp(path.join(path.dirname(directory), "verifier-negative-")), checks = [];
for (const [kind, expected] of [["snapshot-mismatch", "Reopened copy differs"], ["cache-disagreement", "Native saved source/cache direct text differs"], ["consistent-identity-rewrite", "content identity/role/text changed"]]) {
  const target = path.join(temporary, kind); await cp(directory, target, { recursive: true });
  await run("python3", ["-c", String.raw`
from pathlib import Path
import sys,zipfile,xml.etree.ElementTree as E,uuid,shutil
folder=Path(sys.argv[1]);kind=sys.argv[2];saved=folder/'process-baseline-saved.pptx';opened=folder/'process-baseline.pptx'
if kind=='snapshot-mismatch':saved.write_bytes(saved.read_bytes()+b'changed')
else:
 with zipfile.ZipFile(saved) as z:parts={n:z.read(n) for n in z.namelist()}
 if kind=='cache-disagreement':
  name='ppt/diagrams/drawing1.xml';node=E.fromstring(parts[name]);leaf=next(n for n in node.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t') if n.text);leaf.text='Tampered cache text';parts[name]=E.tostring(node,encoding='utf-8',xml_declaration=True)
 else:
  target=uuid.UUID(sys.argv[3]);replacement='{DDEEFF00-1122-3344-5566-778899AABBCC}';name='ppt/diagrams/data1.xml';node=E.fromstring(parts[name]);count=0
  for n in node.iter():
   for k,v in list(n.attrib.items()):
    try:equal=uuid.UUID(v)==target
    except ValueError:equal=False
    if equal:n.set(k,replacement);count+=1
  assert count>=3;parts[name]=E.tostring(node,encoding='utf-8',xml_declaration=True)
 with zipfile.ZipFile(saved,'w',compression=zipfile.ZIP_DEFLATED) as z:
  for n,b in parts.items():z.writestr(n,b)
 shutil.copyfile(saved,opened)
`, target, kind, plan.cases.find(c => c.family === "process" && c.kind === "baseline").selection.sourceModelId], { timeout: 30000 });
  const output = path.join(target, "negative-output.json"); let failure;
  try { await run(process.execPath, [new URL("./inspect-smartart-stale-cache-powerpoint.mjs", import.meta.url).pathname, target, output], { timeout: 60000, maxBuffer: 2 * 1024 * 1024 }); }
  catch (error) { failure = error; }
  assert.ok(failure && failure.code === 1, `Expected explicit validation rejection: ${kind}`);
  assert.ok(failure.stderr.includes(expected), `${kind}: ${failure.stderr}`);
  await assert.rejects(readFile(output), { code: "ENOENT" });
  checks.push({ case: kind, rejected: true, reason: expected });
}
assert.equal(sha(await readFile(path.join(directory, "process-baseline-saved.pptx"))), sha(original));
assert.equal(sha(await readFile(new URL("../docs/testing/smartart-stale-cache-powerpoint-evidence.json", import.meta.url))), sha(evidence));
const result = { profile: "powerpoint-stale-cache-verifier-negatives-v1", directory: temporary, originalNativeSha256: sha(original), originalEvidenceSha256: sha(evidence), checks, originalNativeAndEvidenceUnchanged: true };
await writeFile(path.join(temporary, "results.json"), JSON.stringify(result, null, 2) + "\n");
await writeFile(new URL("../docs/testing/smartart-stale-cache-powerpoint-negative-evidence.json", import.meta.url), JSON.stringify(result, null, 2) + "\n");
console.log(JSON.stringify(result, null, 2));
