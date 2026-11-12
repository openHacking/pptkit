import assert from "node:assert/strict";
import test from "node:test";
import { inspectPptx, importPptx } from "../dist/index.js";
import { structureFixture, encodeStructure } from "./structure-fixture.mjs";
const p14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
function fixture(strict, body, attrs = 'p14:dur="1234" advClick="0" advTm="0"', extra = "") {
  const f = structureFixture(strict);
  f.parts["deck/z.xml"] = f.p("sld", `<p:cSld/><p:transition xmlns:p14="${p14}" xmlns:mc="${mc}" mc:Ignorable="p14" ${attrs}>${body}</p:transition>${extra}`);
  return encodeStructure(f);
}

test("both dialects map explicit fade milliseconds, zero/manual timing and exact no-op bytes", async () => {
  for (const strict of [false, true]) {
    const bytes = fixture(strict, '<p:fade/>'), inspected = await inspectPptx(bytes);
    assert.deepEqual(inspected.slides[0].transition, { status: "mapped", speed: "fast", type: "fade", durationMs: 1234, onClick: false, afterMs: 0 });
    for (const projection of ["structure", "content"]) {
      const s = await importPptx(bytes, { projection });
      assert.deepEqual(s.document.slides[0].transition, { type: "fade", durationMs: 1234, onClick: false, afterMs: 0 });
      const prep = await s.prepareSave({ mode: "no-op" }); assert.equal(prep.status, "ready"); assert.deepEqual((await s.save(prep.plan)).bytes, new Uint8Array(bytes));
      s.document.slides[0].setTransition(); await assert.rejects(s.save(prep.plan), { code: "save-snapshot-stale" }); s.dispose();
    }
  }
});

test("unknown effects, sounds, through-black, speed-only and ignored raw extensions never gain fade permission", async () => {
  for (const [body, attrs] of [
    ['<p:futureEffect/>', 'p14:dur="12"'], ['<p:fade thruBlk="1"/>', 'p14:dur="12"'], ['<p:fade/>', 'spd="slow"'],
    ['<p:fade/><p:sndAc/>', 'p14:dur="12"'], ['<p:fade/><p14:future/>', 'p14:dur="12"'], ['<p:fade p14:future="1"/>', 'p14:dur="12"'],
    ['<p:fade/>', 'p14:dur="12" p14:future="1"'],
  ]) {
    const bytes = fixture(false, body, attrs), s = await importPptx(bytes, { projection: "structure" });
    assert.equal(s.report.inspection.slides[0].transition.status, "source-only"); assert.equal(s.document.slides[0].transition, undefined);
    const prep = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(prep.plan)).bytes, new Uint8Array(bytes)); s.dispose();
  }
});

test("malformed mapped transition scalar fields and duplicate transitions reject with source locations", async () => {
  for (const [body, attrs, extra] of [
    ['<p:fade/>', 'p14:dur="-1"'], ['<p:fade/>', 'p14:dur="1.5"'], ['<p:fade/>', 'p14:dur="4294967296"'],
    ['<p:fade/>', 'advClick="yes"'], ['<p:fade/>', 'advTm="-1"'], ['<p:fade/>', 'spd="super"'],
    ['<p:fade thruBlk="yes"/>', ''], ['<p:fade/><p:fade/>', ''], ['<p:fade/>', '', '<p:transition/>'],
  ]) await assert.rejects(inspectPptx(fixture(false, body, attrs, extra)), e => e.code === "presentation-invalid" && e.entry === "/deck/z.xml");
});


test("MCE fallback keeps its own speed and never borrows duration from an inactive branch", async () => {
  const f = structureFixture();
  f.parts["deck/z.xml"] = f.p("sld", `<p:cSld/><mc:AlternateContent xmlns:mc="${mc}" xmlns:p14="${p14}"><mc:Choice Requires="p14"><p:transition p14:dur="1234"><p:fade/></p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>`);
  const result = await inspectPptx(encodeStructure(f));
  assert.equal(result.slides[0].transition.status, "source-only"); assert.equal(result.slides[0].transition.speed, "slow"); assert.equal(result.slides[0].transition.durationMs, undefined);
});


test("both dialects map every Push/Wipe cardinal direction with exact timing", async () => {
  for (const strict of [false, true]) for (const type of ["push", "wipe"]) for (const [dir, direction] of [["l", "left"], ["r", "right"], ["u", "up"], ["d", "down"]]) {
    const bytes = fixture(strict, `<p:${type} dir="${dir}"/>`);
    const session = await importPptx(bytes, { projection: "content" });
    assert.deepEqual(session.document.slides[0].transition, { type, direction, durationMs: 1234, onClick: false, afterMs: 0 }); session.dispose();
  }
});

test("side transition missing/unknown/opaque direction does not silently acquire editing permission", async () => {
  for (const type of ["push", "wipe"]) {
    await assert.rejects(inspectPptx(fixture(false, `<p:${type} dir="diagonal"/>`)), { code: "presentation-invalid" });
    for (const body of [`<p:${type}/>`, `<p:${type} dir="l" p14:future="1"/>`, `<p:${type} dir="l"><p14:opaque/></p:${type}>`]) assert.equal((await inspectPptx(fixture(false, body))).slides[0].transition.status, "source-only");
  }
});
