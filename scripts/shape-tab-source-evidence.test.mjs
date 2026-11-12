import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

test("shape tab source comparison detects changed text, geometry and fill even with a recomputed hash", async t => {
  const check = args => execFileSync("python3", ["scripts/research/verify-shape-tab-source.py", ...args], { encoding: "utf8", stdio: "pipe" });
  const valid = JSON.parse(check([]));
  assert.equal(valid.items.length, 12);
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-shape-tab-source-"));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const mutate = `
import hashlib, json, sys, zipfile
from pathlib import Path
src, target, manifest, old, new = sys.argv[1:]
with zipfile.ZipFile(src) as inp, zipfile.ZipFile(target, 'w') as out:
    for info in inp.infolist():
        data = inp.read(info.filename)
        if info.filename == 'ppt/slides/slide1.xml':
            start = data.index(b'name="left-1"')
            end = data.index(b'</p:sp>', start)
            target_shape = data[start:end]
            assert old.encode() in target_shape
            data = data[:start] + target_shape.replace(old.encode(), new.encode(), 1) + data[end:]
        out.writestr(info, data)
record = json.loads(Path('.tmp/native-tab-stops/powerpoint-shape-j8Dn0A/manifest.json').read_text())
record['sha256'] = hashlib.sha256(Path(target).read_bytes()).hexdigest()
Path(manifest).write_text(json.dumps(record))
`;
  for (const [oldValue, newValue] of [[">WIDE<", ">WIDF<"], ['x="762000"', 'x="762001"'], ['val="F2F5FA"', 'val="F2F5FB"']]) {
    const output = path.join(dir, "changed.pptx"), manifest = path.join(dir, "manifest.json");
    execFileSync("python3", ["-c", mutate, ".tmp/native-tab-stops/powerpoint-shape-j8Dn0A/shapes.pptx", output, manifest, oldValue, newValue]);
    assert.throws(() => check(["--shape", output, "--manifest", manifest]));
  }
});
