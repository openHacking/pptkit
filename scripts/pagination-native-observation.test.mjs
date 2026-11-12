import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { cp, mkdtemp, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

test("native pagination observation detects altered package and visible pixels", async t => {
  const root = ".tmp/pagination-native-observation";
  try { await stat(path.join(root, "style-probe.pdf")); } catch { t.skip("Retained native observation files are unavailable"); return; }
  const verify = directory => spawnSync("python3", ["scripts/research/inspect-pagination-native.py", directory], { encoding: "utf8" });
  const valid = verify(root); assert.equal(valid.status, 0, valid.stderr);
  const temp = await mkdtemp(path.join(tmpdir(), "pptkit-pagination-evidence-"));
  t.after(() => rm(temp, { recursive: true, force: true }));
  for (const name of ["source.pptx", "style-probe.pptx", "style-probe.pdf", "pptkit-pagination-probe-20261111.pdf"]) await cp(path.join(root, name), path.join(temp, name));
  const mutate = (python, ...args) => {
    const run = spawnSync("python3", ["-c", python, ...args], { encoding: "utf8" });
    assert.equal(run.status, 0, run.stderr);
  };
  mutate('import fitz,sys; p=fitz.open(sys.argv[1]); p[0].draw_rect(fitz.Rect(158,78,162,82), color=(0,0,0), fill=(0,0,0)); p.save(sys.argv[2]); p.close()', path.join(temp, "style-probe.pdf"), path.join(temp, "corrupt.pdf"));
  await cp(path.join(temp, "corrupt.pdf"), path.join(temp, "style-probe.pdf"));
  assert.notEqual(verify(temp).status, 0, "a changed sampled pixel must fail");
  await cp(path.join(root, "style-probe.pdf"), path.join(temp, "style-probe.pdf"));
  mutate('import sys,zipfile; a=zipfile.ZipFile(sys.argv[1]); b=zipfile.ZipFile(sys.argv[2],"w"); [(b.writestr(i,a.read(i).replace(b\'lastRow="1"\',b\'lastRow="0"\') if i.filename=="ppt/slides/slide1.xml" else a.read(i))) for i in a.infolist()]; b.close(); a.close()', path.join(temp, "style-probe.pptx"), path.join(temp, "corrupt.pptx"));
  await cp(path.join(temp, "corrupt.pptx"), path.join(temp, "style-probe.pptx"));
  assert.notEqual(verify(temp).status, 0, "a changed native switch must fail");
});

test("source-boundary observation rejects changed row flags and sampled colors", async t => {
  const root = ".tmp/pagination-native-observation/source-boundaries";
  try { await stat(path.join(root, "style-probe-boundaries.pdf")); } catch { t.skip("Retained native observation files are unavailable"); return; }
  const verify = directory => spawnSync("python3", ["scripts/research/inspect-pagination-boundaries.py", directory], { encoding: "utf8" });
  assert.equal(verify(root).status, 0);
  const temp = await mkdtemp(path.join(tmpdir(), "pptkit-boundary-evidence-"));
  t.after(() => rm(temp, { recursive: true, force: true }));
  for (const name of ["source.pptx", "style-probe-boundaries.pptx", "style-probe-boundaries.pdf"]) await cp(path.join(root, name), path.join(temp, name));
  const pdf = path.join(temp, "style-probe-boundaries.pdf"), corrupt = path.join(temp, "corrupt.pdf");
  const mutation = spawnSync("python3", ["-c", 'import fitz,sys; p=fitz.open(sys.argv[1]); p[1].draw_rect(fitz.Rect(158,118,162,122), color=(0,0,0), fill=(0,0,0)); p.save(sys.argv[2]); p.close()', pdf, corrupt], { encoding: "utf8" });
  assert.equal(mutation.status, 0, mutation.stderr);
  await cp(corrupt, pdf);
  assert.notEqual(verify(temp).status, 0);
  await cp(path.join(root, "style-probe-boundaries.pdf"), pdf);
  const bad = path.join(temp, "bad.pptx"), pptx = path.join(temp, "style-probe-boundaries.pptx");
  const change = spawnSync("python3", ["-c", 'import sys,zipfile; a=zipfile.ZipFile(sys.argv[1]); b=zipfile.ZipFile(sys.argv[2],"w"); [(b.writestr(i,a.read(i).replace(b\'firstRow="1"\',b\'firstRow="0"\') if i.filename=="ppt/slides/slide2.xml" else a.read(i))) for i in a.infolist()]; b.close(); a.close()', pptx, bad], { encoding: "utf8" });
  assert.equal(change.status, 0, change.stderr);
  await cp(bad, pptx);
  assert.notEqual(verify(temp).status, 0);
});

test("saved pagination observation rejects changed text, row switches and reopened pixels", async t => {
  const root = ".tmp/pagination-native-observation/source-boundaries";
  try { await stat(path.join(root, "native-reopened.pdf")); } catch { t.skip("Retained PowerPoint save/reopen files are unavailable"); return; }
  const verify = directory => spawnSync("python3", ["scripts/research/inspect-pagination-reopen.py", directory], { encoding: "utf8" });
  const valid = verify(root); assert.equal(valid.status, 0, valid.stderr);
  const temp = await mkdtemp(path.join(tmpdir(), "pptkit-pagination-reopen-"));
  t.after(() => rm(temp, { recursive: true, force: true }));
  for (const name of ["source.pptx", "native-saved.pptx", "native-saved.pdf", "native-reopened.pdf"]) await cp(path.join(root, name), path.join(temp, name));
  const saved = path.join(temp, "native-saved.pptx"), corrupt = path.join(temp, "corrupt.pptx");
  for (const [oldValue, newValue] of [['ROW 2', 'ROW X'], ['firstRow="1"', 'firstRow="0"']]) {
    const mutation = spawnSync("python3", ["-c", 'import sys,zipfile; a=zipfile.ZipFile(sys.argv[1]); b=zipfile.ZipFile(sys.argv[2],"w"); [(b.writestr(i,a.read(i).replace(sys.argv[3].encode(),sys.argv[4].encode()) if i.filename=="ppt/slides/slide3.xml" else a.read(i))) for i in a.infolist()]; b.close(); a.close()', saved, corrupt, oldValue, newValue], { encoding: "utf8" });
    assert.equal(mutation.status, 0, mutation.stderr); await cp(corrupt, saved);
    assert.notEqual(verify(temp).status, 0, oldValue);
    await cp(path.join(root, "native-saved.pptx"), saved);
  }
  const pdf = path.join(temp, "native-reopened.pdf"), badPdf = path.join(temp, "corrupt.pdf");
  const mutation = spawnSync("python3", ["-c", 'import fitz,sys; p=fitz.open(sys.argv[1]); p[2].draw_rect(fitz.Rect(158,118,162,122),color=(0,0,0),fill=(0,0,0)); p.save(sys.argv[2]); p.close()', pdf, badPdf], { encoding: "utf8" });
  assert.equal(mutation.status, 0, mutation.stderr); await cp(badPdf, pdf);
  assert.notEqual(verify(temp).status, 0, "changed reopened page must fail");
});
