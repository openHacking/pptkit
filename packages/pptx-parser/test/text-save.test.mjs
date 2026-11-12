import assert from "node:assert/strict";
import test from "node:test";
import { createHash } from "node:crypto";
import { createPresentation, normalizePresentation } from "../../core/dist/index.js";
import { importPptx, compareOpcPackages, DEFAULT_SESSION_LIMITS, inspectOpcClosure } from "../dist/index.js";
import { analyzeTextEdits } from "../dist/import/analyze-text-edits.js";
import { contentFixture, shape, body } from "./content-fixture.mjs";
import { imageFixture, encodeImages, picture } from "./image-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";
import { readFile } from "node:fs/promises";
import { encodeStructure } from "./structure-fixture.mjs";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";
import { zip } from "./zip-fixture.mjs";
import { getEventListeners } from "node:events";

const profile = { projection: "content", editSaveProfile: "text-runs-v1" };
const code = expected => error => { assert.equal(error.code, expected); return true; };
const digest = bytes => createHash("sha256").update(bytes).digest("hex");
const repack = parts => zip([...parts].map(([name, bytes]) => ({ name, body: bytes }))).bytes;
const fixture = (strict = false) => contentFixture(shape({ txBox: true, text: body(), props: '<a:effectLst><a:glow rad="12700"><a:srgbClr val="112233"/></a:glow></a:effectLst>' }), strict);
function replaceElement(session, modify, index = 0) {
  const slide = session.document.slides[0], old = slide.removeElement(slide.elements[index].id), next = structuredClone(old);
  modify(next); slide.insertElement(index, next); return old;
}
const changeText = (session, text) => replaceElement(session, e => { e.content[0].runs[0].text = text; });
async function save(session) {
  const prepared = await session.prepareSave({ mode: "edit" });
  assert.equal(prepared.status, "ready", JSON.stringify(prepared.report));
  const result = await session.save(prepared.plan);
  return { ...result, plan: prepared.plan };
}
const freezeCheck = value => {
  if (!value || typeof value !== "object") return;
  assert.ok(Object.isFrozen(value));
  for (const child of Object.values(value)) freezeCheck(child);
};

test("explicit profile binds raw runs and performs loss-free text saves in both dialects", async () => {
  for (const strict of [false, true]) {
    const f = fixture(strict), original = encodeStructure(f), session = await importPptx(original, profile);
    try {
      assert.equal(session.report.textRuns.length, 1);
      const binding = session.report.textRuns[0]; assert.equal(binding.status, "bound");
      assert.equal(binding.source.elementIndex, binding.sourceElementIndex); freezeCheck(session.report.textRuns);
      assert.equal(Object.hasOwn(binding, 'rawRunElementIndex'), false);
      assert.equal(Object.hasOwn(binding, 'rawParagraphElementIndex'), false);
      assert.equal((await session.prepareSave({ mode: "edit" })).report.diagnostics[0].code, "save-edit-no-changes");
      changeText(session, "保存 <&>🙂\r\n");
      const result = await save(session); freezeCheck(result.report);
      assert.equal(result.report.assurance, "scoped-preservation"); assert.deepEqual(result.report.losses, []);
      assert.deepEqual(result.report.changes.map(c => [c.kind, c.documentPath, c.text]), [["text", "/slides/0/elements/0/content/0/runs/0/text", { before: " A & B ", after: "保存 <&>🙂\r\n" }]]);
      assert.equal(result.report.changes[0].source.elementIndex, binding.sourceElementIndex);
      const parts = readGeneratedZip(result.bytes), before = readGeneratedZip(original);
      for (const [name, payload] of parts) if (name !== "deck/z.xml") assert.equal(digest(payload), digest(before.get(name)), name);
      assert.equal(parts.get("deck/z.xml").toString(), f.parts["deck/z.xml"].replace("> A &amp; B </a:t>", ">保存 &lt;&amp;&gt;🙂&#xD;\n</a:t>"));
      const reopened = await importPptx(result.bytes, profile);
      assert.deepEqual(normalizePresentation(reopened.document), normalizePresentation(session.document));
      const noop = await reopened.prepareSave({ mode: "no-op" }); assert.equal(noop.status, "ready");
      assert.equal(digest((await reopened.save(noop.plan)).bytes), digest(result.bytes));
      reopened.dispose();
      assert.equal((await session.prepareSave({ mode: "no-op" })).status, "blocked");
      result.plan.dispose();
    } finally { session.dispose(); }
  }
});

test("text-only source saves retain shared and opaque image resources while image edits remain blocked", async () => {
  const wmf = await readFile(new URL('../../../tests/fixtures/metafile-libreoffice-asymmetric.wmf', import.meta.url));
  const svg = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"><rect width="8" height="4" fill="red"/></svg>');
  for (const strict of [false, true]) {
    const extension = '<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><v:svgBlip xmlns:v="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="svg"/></a:ext></a:extLst>';
    const text = shape({ id: '2', txBox: true, text: body() });
    const f = imageFixture(text + picture({ id: 3, effects: extension }) + picture({ id: 5 }) + picture({ id: 4, embed: 'legacy' }), strict);
    f.parts['media/vector.svg'] = svg; f.mediaTypes['media/vector.svg'] = 'image/svg+xml';
    f.parts['media/legacy.wmf'] = wmf; f.mediaTypes['media/legacy.wmf'] = 'image/wmf';
    f.parts['deck/_rels/z.xml.rels'] = relationships([f.rel('layout','slideLayout','layout.xml'),f.rel('image','image','../media/raster.bin'),f.rel('svg','image','../media/vector.svg'),f.rel('legacy','image','../media/legacy.wmf')]);
    const original=encodeImages(f), session=await importPptx(original,profile);
    try {
      assert.equal(session.document.assets.length,1);assert.equal(session.document.slides[0].elements.filter(item=>item.type==='image').length,2);assert.ok(session.report.findings.some(item=>/SVG image extension/.test(item.reason)));
      changeText(session,'changed');const saved=await save(session),before=readGeneratedZip(original),after=readGeneratedZip(saved.bytes);
      for(const [name,bytes] of before)if(name!=='deck/z.xml')assert.deepEqual(after.get(name),bytes,name);
      assert.equal(after.get('deck/z.xml').toString(),f.parts['deck/z.xml'].replace('> A &amp; B </a:t>','>changed</a:t>'));
      let current=saved.bytes;
      for(const [prior,next] of [['changed','changed-2'],['changed-2','changed-3']]){
        const round=await importPptx(current,profile);try{
          changeText(round,next);const output=await save(round),oldParts=readGeneratedZip(current),newParts=readGeneratedZip(output.bytes);
          for(const [name,bytes] of oldParts)if(name!=='deck/z.xml')assert.deepEqual(newParts.get(name),bytes,name);
          assert.equal(newParts.get('deck/z.xml').toString(),oldParts.get('deck/z.xml').toString().replace(`>${prior}</a:t>`,`>${next}</a:t>`));
          current=output.bytes;
        }finally{round.dispose()}
      }
      const reopened=await importPptx(current,profile);try{
        const target=reopened.document.slides[0].elements.find(item=>item.type==='image'),changed=structuredClone(target);changed.box.x+=1;
        reopened.document.slides[0].replaceElement(target.id,changed);
        assert.equal((await reopened.prepareSave({mode:'edit'})).status,'blocked');
        reopened.document.slides[0].replaceElement(target.id,target);
        reopened.document.registerAsset({id:'replacement',kind:'image',source:{type:'bytes',value:f.parts['media/raster.bin']},mimeType:'image/png'});
        const replaced=structuredClone(target);replaced.assetId='replacement';reopened.document.slides[0].replaceElement(target.id,replaced);
        assert.equal((await reopened.prepareSave({mode:'edit'})).status,'blocked');
        changeText(reopened,'other text');assert.equal((await reopened.prepareSave({mode:'edit'})).status,'blocked');
      }finally{reopened.dispose()}
    }finally{session.dispose()}
  }
});

test("default sessions keep their existing edit/rebuild behavior and reject invalid profile combinations", async () => {
  const bytes = encodeStructure(fixture()), session = await importPptx(bytes, { projection: "content" });
  changeText(session, "new");
  for (const mode of ["edit", "rebuild"]) assert.equal((await session.prepareSave({ mode })).report.diagnostics[0].code, "save-mode-unsupported");
  assert.equal(session.report.textRuns, undefined); session.dispose();
  await assert.rejects(importPptx(bytes, { projection: "structure", editSaveProfile: "text-runs-v1" }), TypeError);
  await assert.rejects(importPptx(bytes, { projection: "content", editSaveProfile: "unknown" }), TypeError);
});

test("MCE raw origin mapping edits only the selected duplicate and retains inactive content", async () => {
  const f = fixture(), selected = shape({ txBox: true, text: body(), props: '<a:effectLst><a:glow rad="12700"><a:srgbClr val="112233"/></a:glow></a:effectLst>' }), hidden = selected.replace("> A &amp; B </a:t>", ">HIDDEN</a:t>");
  f.parts["deck/z.xml"] = f.parts["deck/z.xml"]
    .replace("<p:sld ", '<p:sld xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ')
    .replace(selected, `<mc:AlternateContent><mc:Choice Requires="x">${hidden}</mc:Choice><mc:Fallback>${selected}</mc:Fallback></mc:AlternateContent>`);
  const input = encodeStructure(f), session = await importPptx(input, profile);
  changeText(session, "selected");
  const result = await save(session), xml = readGeneratedZip(result.bytes).get("deck/z.xml").toString();
  assert.equal(xml, f.parts["deck/z.xml"].replace("> A &amp; B </a:t>", ">selected</a:t>"));
  assert.ok(xml.includes(">HIDDEN</a:t>"));
  assert.ok(session.report.textRuns[0].sourceElementIndex > 40);
  session.dispose();
});

test("repeated text targets, multiple runs and empty/CDATA character atoms retain source order and styles", async () => {
  const paragraphs = '<a:p><a:r><a:rPr b="1"/><a:t><![CDATA[same]]></a:t></a:r><a:r><a:rPr i="1"/><a:t>same</a:t></a:r><a:r><a:t /></a:r></a:p>';
  const f = contentFixture(shape({ txBox: true, text: body(paragraphs) })), session = await importPptx(encodeStructure(f), profile);
  replaceElement(session, e => { e.content[0].runs[1].text = "second"; e.content[0].runs[2].text = "empty filled"; });
  const result = await save(session), xml = readGeneratedZip(result.bytes).get("deck/z.xml").toString();
  assert.equal(xml, f.parts["deck/z.xml"].replace('>same</a:t>', '>second</a:t>').replace('<a:t />', '<a:t >empty filled</a:t>'));
  assert.equal(result.report.changes.length, 2);
  assert.ok(xml.includes("<![CDATA[same]]>")); session.dispose();
});

test("mixed-source, break and synthetic runs remain unmapped and refuse edits atomically", async () => {
  for (const paragraphs of [
    '<a:p><a:r><a:t> A <!--keep-->&amp; B </a:t></a:r></a:p>',
    '<a:p><a:br/></a:p>', '<a:p><a:endParaRPr b="1"/></a:p>',
  ]) {
    const session = await importPptx(encodeStructure(contentFixture(shape({ txBox: true, text: body(paragraphs) }))), profile);
    assert.equal(session.report.textRuns[0].status, "unmapped");
    changeText(session, "replacement");
    const result = await session.prepareSave({ mode: "edit" });
    assert.equal(result.status, "blocked"); assert.equal(result.report.assurance, "none"); assert.equal(result.report.changes[0].kind, "text");
    session.dispose();
  }
});

test("text-subtarget/opaque animation references are guarded across all branches and normalized identities", async () => {
  for (const target of ['<p:spTgt spid="2"><p:txEl><p:charRg st="0" end="2"/></p:txEl></p:spTgt>', '<p:spTgt spid=" +0002 "><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt>', '<p:spTgt spid="2" custom="opaque"/>', '<p:spTgt spid="bad"/>']) {
    const f = fixture();
    f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace("</p:sld>", `<p:timing>${target}</p:timing></p:sld>`);
    const session = await importPptx(encodeStructure(f), profile); changeText(session, "new");
    const result = await session.prepareSave({ mode: "edit" });
    assert.equal(result.status, "blocked");
    assert.equal(result.report.diagnostics[0].code, "save-text-range-reference"); assert.ok(result.report.diagnostics[0].source.elementIndex);
    session.dispose();
  }
  const f = fixture();
  f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace("<p:sld ", '<p:sld xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ')
    .replace("</p:sld>", '<mc:AlternateContent><mc:Choice Requires="x"><p:timing><p:spTgt spid="2"><p:txEl/></p:spTgt></p:timing></mc:Choice><mc:Fallback/></mc:AlternateContent></p:sld>');
  const session = await importPptx(encodeStructure(f), profile); changeText(session, "new");
  assert.equal((await session.prepareSave({ mode: "edit" })).report.diagnostics[0].code, "save-text-range-reference"); session.dispose();
});

test("whole-shape targets and ranges owned by a different object retain their exact source XML", async () => {
  for (const target of ['<p:spTgt spid="2"/>', '<p:spTgt spid="3"><p:txEl/></p:spTgt>']) {
    const f = fixture(); f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace("</p:sld>", `<p:timing>${target}</p:timing></p:sld>`);
    const session = await importPptx(encodeStructure(f), profile); changeText(session, "new");
    const result = await save(session);
    assert.ok(readGeneratedZip(result.bytes).get("deck/z.xml").toString().includes(target)); session.dispose();
  }
});

test("style, geometry, asset, topology, notes, metadata and custom-data mutations are classified and blocked", async () => {
  const modifications = [
    ["style", s => replaceElement(s, e => { e.content[0].runs[0].style.bold = false; })],
    ["geometry", s => replaceElement(s, e => { e.box.x += 10; })],
    ["structure", s => s.document.duplicateSlide(s.document.slides[0].id)],
    ["structure", s => replaceElement(s, e => { e.content[0].runs.push({ text: "added" }); })],
    ["text", s => s.document.slides[0].setNotes("new notes")],
    ["asset", s => s.document.registerAsset({ id: "new", kind: "image", source: { type: "bytes", value: Uint8Array.of(1, 2, 3) } })],
  ];
  for (const [kind, modify] of modifications) {
    const session = await importPptx(encodeStructure(fixture()), profile); modify(session);
    const result = await session.prepareSave({ mode: "edit" }); assert.equal(result.status, "blocked");
    assert.ok(result.report.changes.some(c => c.kind === kind), JSON.stringify(result.report.changes));
    freezeCheck(result.report); session.dispose();
  }
  // Custom data deliberately resembles normalized text: it cannot receive the
  // real text element's derived-plainText exemption.
  const doc = createPresentation(); doc.addSlide({ id: "s", customData: { record: { type: "text", content: [], plainText: "old", size: 1 } } });
  const before = normalizePresentation(doc), after = structuredClone(before);
  after.slides[0].customData.record.plainText = "new";
  const result = await analyzeTextEdits(before, after, [], [], DEFAULT_SESSION_LIMITS, async () => {});
  assert.deepEqual(result.changes.map(c => c.kind), ["other"]); assert.equal(result.diagnostics.length, 1);
  const metadata = structuredClone(before); metadata.metadata.title = "new title";
  assert.deepEqual((await analyzeTextEdits(before, metadata, [], [], DEFAULT_SESSION_LIMITS, async () => {})).changes.map(c => c.kind), ["metadata"]);
});

test("fixed-topology profile blocks visually unchanged splits and duplicate-run merges atomically", async () => {
  for (const strict of [false, true]) {
    const text = body('<a:p><a:r><a:t>e\u0301🙂</a:t></a:r><a:r><a:t>same</a:t></a:r><a:r><a:t>same</a:t></a:r></a:p>');
    const original = encodeStructure(contentFixture(shape({ txBox: true, text }), strict));
    const session = await importPptx(original, profile);
    try {
      const baseline = structuredClone(session.document.slides[0].elements[0]);
      for (const split of [1, 2, 3]) {
        replaceElement(session, element => {
          const run = element.content[0].runs[0];
          element.content[0].runs.splice(0, 1, { ...run, text: run.text.slice(0, split) }, { ...run, text: run.text.slice(split) });
        });
        const blocked = await session.prepareSave({ mode: 'edit' });
        assert.equal(blocked.status, 'blocked');
        assert.ok(blocked.report.diagnostics.some(item => item.code === 'save-edit-unsupported'));
        replaceElement(session, element => Object.assign(element, structuredClone(baseline)));
      }
      replaceElement(session, element => {
        const runs = element.content[0].runs;
        runs.splice(1, 2, { ...runs[1], text: runs[1].text + runs[2].text });
      });
      assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked');
      replaceElement(session, element => Object.assign(element, structuredClone(baseline)));
      const noop = await session.prepareSave({ mode: 'no-op' });
      assert.equal(noop.status, 'ready');
      assert.equal(digest((await session.save(noop.plan)).bytes), digest(original));
    } finally { session.dispose(); }
  }
});

test("pre-split source runs retain distinct bindings and formatting through three scalar saves", async () => {
  const props = '<a:rPr b="1" sz="2400"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:rPr>';
  const paragraph = `<a:p><a:pPr algn="ctr"/><a:r>${props}<a:t>e\u0301</a:t></a:r><a:r>${props}<a:t>🙂</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>`;
  for (const strict of [false, true]) {
    const f = contentFixture(shape({ txBox: true, text: body(paragraph) }), strict);
    let bytes = encodeStructure(f);
    for (let round = 1; round <= 3; round++) {
      const session = await importPptx(bytes, profile);
      try {
        const runs = session.document.slides[0].elements[0].content[0].runs;
        assert.equal(runs.length, 2);
        assert.equal(runs[0].style.bold, true); assert.equal(runs[1].style.bold, true);
        assert.equal(session.report.textRuns.length, 2);
        assert.ok(session.report.textRuns.every(binding => binding.status === 'bound'));
        assert.notEqual(session.report.textRuns[0].sourceElementIndex, session.report.textRuns[1].sourceElementIndex);
        const before = readGeneratedZip(bytes), old = runs.map(run => run.text), next = [`L${round}`, `R${round}`];
        replaceElement(session, element => { element.content[0].runs[0].text = next[0]; element.content[0].runs[1].text = next[1]; });
        const result = await save(session), after = readGeneratedZip(result.bytes);
        for (const [name, payload] of before) if (name !== 'deck/z.xml') assert.deepEqual(after.get(name), payload, name);
        const expected = before.get('deck/z.xml').toString().replace(`>${old[0]}</a:t>`, `>${next[0]}</a:t>`).replace(`>${old[1]}</a:t>`, `>${next[1]}</a:t>`);
        assert.equal(after.get('deck/z.xml').toString(), expected);
        result.plan.dispose(); bytes = result.bytes;
      } finally { session.dispose(); }
    }
  }
});

test("same-value and restored Core state use original-byte no-op; edit saves never advance the baseline", async () => {
  const input = encodeStructure(fixture()), session = await importPptx(input, profile);
  const old = changeText(session, "first"); const first = await save(session); first.plan.dispose();
  changeText(session, "second"); const second = await save(session);
  assert.deepEqual(second.report.changes[0].text, { before: " A & B ", after: "second" });
  assert.ok(readGeneratedZip(second.bytes).get("deck/z.xml").toString().includes(">second</a:t>"));
  replaceElement(session, e => { Object.assign(e, structuredClone(old)); });
  const noop = await session.prepareSave({ mode: "no-op" }); assert.equal(noop.status, "ready");
  assert.equal(digest((await session.save(noop.plan)).bytes), digest(input));
  assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked"); session.dispose();
});

test("stale, foreign/disposed plans, exact consent and detached output retain session guarantees", async () => {
  const input = encodeStructure(fixture()), a = await importPptx(input, profile), b = await importPptx(input, profile);
  changeText(a, "new"); const saved = await save(a);
  await assert.rejects(b.save(saved.plan), code("plan-session-mismatch"));
  await assert.rejects(a.save(saved.plan, { acceptLossIds: ["anything"] }), code("save-loss-consent-mismatch"));
  const hash = digest(saved.bytes); saved.bytes.fill(0);
  assert.equal(digest((await a.save(saved.plan)).bytes), hash);
  changeText(a, "changed again");
  await assert.rejects(a.save(saved.plan), code("save-snapshot-stale"));
  changeText(a, "new"); assert.equal(digest((await a.save(saved.plan)).bytes), hash);
  saved.plan.dispose(); await assert.rejects(a.save(saved.plan), code("plan-disposed"));
  a.dispose(); b.dispose();
});

test("independent analysis reconciles shared raw targets and ignores only real derived plainText", async () => {
  const doc = createPresentation(), slide = doc.addSlide({ id: "s" });
  for (const id of ["a", "b"]) slide.addElement({ id, type: "text", content: [{ runs: [{ text: "old" }] }], box: { x: 0, y: 0, width: 100, height: 30 } });
  const baseline = normalizePresentation(doc), candidate = structuredClone(baseline);
  candidate.slides[0].elements[0].content[0].runs[0].text = "new"; candidate.slides[0].elements[0].plainText = "new";
  const bindings = ["a", "b"].map(elementId => ({ slideId: "s", elementId, paragraphIndex: 0, runIndex: 0, source: { part: "/shared.xml", sourceObjectId: "2" }, sourceElementIndex: 3, status: "bound", baselineText: "old", namespaceUri: "urn:a" }));
  const conflict = await analyzeTextEdits(baseline, candidate, bindings, [], DEFAULT_SESSION_LIMITS, async () => {});
  assert.equal(conflict.changes.length, 1); assert.equal(conflict.diagnostics[0].code, "save-source-shared-conflict");
  candidate.slides[0].elements[1].content[0].runs[0].text = "new"; candidate.slides[0].elements[1].plainText = "new";
  const agreed = await analyzeTextEdits(baseline, candidate, bindings, [], DEFAULT_SESSION_LIMITS, async () => {});
  assert.deepEqual(agreed.diagnostics, []); assert.equal(agreed.changes.length, 2); assert.equal(agreed.parts.get("/shared.xml").length, 1);
});

test("complete change/work/validation/output budgets and cancellations release preparation reservations", async () => {
  const input = encodeStructure(fixture());
  for (const sessionLimits of [{ maxSaveAnalysisWork: 1 }, { maxSaveValidationBytes: 1 }, { maxSaveValidationWork: 1 }, { maxPreparedOutputBytes: 100 }]) {
    const session = await importPptx(input, { ...profile, sessionLimits }); changeText(session, "new");
    await assert.rejects(session.prepareSave({ mode: "edit" }), code("save-limit-exceeded"));
    session.dispose();
  }
  const f = contentFixture(shape({ txBox: true, text: body('<a:p><a:r><a:t>a</a:t></a:r><a:r><a:t>b</a:t></a:r></a:p>') }));
  const session = await importPptx(encodeStructure(f), { ...profile, sessionLimits: { maxSaveChanges: 1, maxLivePlans: 1 } });
  replaceElement(session, e => { e.content[0].runs[0].text = "x"; e.content[0].runs[1].text = "y"; });
  await assert.rejects(session.prepareSave({ mode: "edit" }), code("save-limit-exceeded"));
  replaceElement(session, e => { e.content[0].runs[1].text = "b"; });
  const abort = new AbortController(), pending = session.prepareSave({ mode: "edit", signal: abort.signal }); abort.abort();
  await assert.rejects(pending, code("operation-aborted"));
  const ready = await save(session); ready.plan.dispose(); session.dispose();
});

test("concurrent reservations shrink to actual output and failures cannot leak live-plan capacity", async () => {
  const input = encodeStructure(fixture()), maxPreparedOutputBytes = 65536;
  const session = await importPptx(input, { ...profile, sessionLimits: { maxPreparedOutputBytes, maxRetainedOutputBytes: maxPreparedOutputBytes + input.length * 3, maxLivePlans: 4 } });
  changeText(session, "new");
  const pending = [session.prepareSave({ mode: "edit" }), session.prepareSave({ mode: "edit" })], results = await Promise.allSettled(pending);
  assert.equal(results.filter(r => r.status === "fulfilled" && r.value.status === "ready").length, 1);
  assert.equal(results.filter(r => r.status === "rejected" && r.reason.code === "save-limit-exceeded").length, 1);
  const first = results.find(r => r.status === "fulfilled").value.plan, second = await session.prepareSave({ mode: "edit" });
  assert.equal(second.status, "ready"); assert.ok(first.byteLength < maxPreparedOutputBytes);
  first.dispose(); second.plan.dispose(); session.dispose();
});

test("many concurrent preparations share one caller abort relay and release every listener", async () => {
  const input = encodeStructure(fixture()), controller = new AbortController();
  const session = await importPptx(input, { projection: "content", sessionLimits: { maxLivePlans: 32, maxRetainedOutputBytes: input.length * 32 } });
  const pending = Array.from({ length: 24 }, () => session.prepareSave({ mode: "no-op", signal: controller.signal }));
  await new Promise(resolve => setTimeout(resolve, 0));
  assert.ok(getEventListeners(controller.signal, "abort").length <= 1);
  const prepared = await Promise.all(pending); assert.ok(prepared.every(p => p.status === "ready"));
  assert.equal(getEventListeners(controller.signal, "abort").length, 0);
  for (const plan of prepared) plan.plan.dispose();
  session.dispose();
});

test("cancellation and disposal after child subscription release the relay and reserved output", async () => {
  const input = encodeStructure(fixture());
  for (const action of ["abort", "dispose"]) {
    const controller = new AbortController();
    let observeSubscription;
    const subscribed = new Promise(resolve => { observeSubscription = resolve; });
    const add = controller.signal.addEventListener.bind(controller.signal);
    controller.signal.addEventListener = (...args) => {
      add(...args);
      if (args[0] === "abort") observeSubscription();
    };
    const session = await importPptx(input, {
      ...profile,
      sessionLimits: { maxLivePlans: 1, maxPreparedOutputBytes: 65536, maxRetainedOutputBytes: 65536 },
    });
    try {
      changeText(session, "cancel an already reserved text save");
      const pending = session.prepareSave({ mode: "edit", signal: controller.signal });
      const rejected = assert.rejects(pending, code(action === "dispose" ? "session-disposed" : "operation-aborted"));
      // The relay is installed only after the initial snapshot yield and output
      // reservation. This observes child work rather than cancelling at entry.
      await subscribed;
      assert.equal(getEventListeners(controller.signal, "abort").length, 1);
      if (action === "dispose") {
        session.dispose();
        assert.equal(getEventListeners(controller.signal, "abort").length, 0);
      } else controller.abort();
      await rejected;
      assert.equal(getEventListeners(controller.signal, "abort").length, 0);
      if (action === "abort") {
        await assert.rejects(session.prepareSave({ mode: "edit", signal: controller.signal }), code("operation-aborted"));
        const retried = await save(session);
        assert.equal(retried.report.changes[0].text.after, "cancel an already reserved text save");
        retried.plan.dispose();
        const another = await save(session);
        another.plan.dispose();
      }
      assert.equal(digest(input), digest(encodeStructure(fixture())));
    } finally { session.dispose(); }
  }
});

test("disposal during preparation and edits during validation cannot commit or corrupt the source", async () => {
  const input = encodeStructure(fixture()), session = await importPptx(input, profile); changeText(session, "x".repeat(10000));
  const pending = session.prepareSave({ mode: "edit" }); setTimeout(() => session.dispose(), 0);
  await assert.rejects(pending, code("session-disposed"));
  const other = await importPptx(input, profile); changeText(other, "first");
  const editing = other.prepareSave({ mode: "edit" }); setTimeout(() => changeText(other, "second"), 0);
  const result = await editing; assert.equal(result.status, "blocked"); assert.equal(result.report.diagnostics[0].code, "save-snapshot-stale");
  const retry = await save(other); assert.equal(retry.report.changes[0].text.after, "second"); other.dispose();
});

test("invalid XML replacement characters are blocked without poisoning later valid edits", async () => {
  const session = await importPptx(encodeStructure(fixture()), profile); changeText(session, "\0");
  const result = await session.prepareSave({ mode: "edit" }); assert.equal(result.status, "blocked");
  assert.equal(result.report.diagnostics[0].code, "save-text-candidate-unsupported");
  changeText(session, "valid retry"); await save(session); session.dispose();
});

test("inactive missing resources and signature/VBA relationships cannot bypass edit admission", async () => {
  for (const type of ["urn:resource", "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature", "http://schemas.microsoft.com/office/2006/relationships/vbaProject"]) {
    const f = fixture();
    f.parts["deck/_rels/z.xml.rels"] = f.parts["deck/_rels/z.xml.rels"]
      .replace("<Relationships ", '<Relationships xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ')
      .replace("</Relationships>", `<mc:AlternateContent><mc:Choice Requires="x"><Relationship Id="hidden" Type="${type}" Target="missing.bin"/></mc:Choice><mc:Fallback/></mc:AlternateContent></Relationships>`);
    const session = await importPptx(encodeStructure(f), profile); changeText(session, "new");
    const result = await session.prepareSave({ mode: "edit" }); assert.equal(result.status, "blocked");
    assert.equal(result.report.diagnostics[0].code, type === "urn:resource" ? "save-source-closure-incomplete" : "save-text-package-policy");
    session.dispose();
  }
});

test("three edit-save/reopen rounds retain unknown XML/parts, source identities and graph without growth", async () => {
  const f = fixture();
  f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace("</p:sld>", '<p:extLst><p:ext uri="urn:future"><x:opaque xmlns:x="urn:future" ref="2"> A &amp; B </x:opaque></p:ext></p:extLst></p:sld>');
  const parts = readGeneratedZip(encodeStructure(f)); parts.set("orphan.xml", Buffer.from('<opaque xmlns="urn:future"><keep/></opaque>'));
  let bytes = repack(parts), previousSize = 0;
  const original = readGeneratedZip(bytes), closure = await inspectOpcClosure(bytes, { roots: ["/"] });
  for (let round = 1; round <= 3; round++) {
    const session = await importPptx(bytes, profile); changeText(session, `round ${round}`);
    const saved = await save(session), current = readGeneratedZip(saved.bytes), compared = await compareOpcPackages(bytes, saved.bytes);
    assert.deepEqual([...current.keys()].sort(), [...original.keys()].sort());
    for (const [name, payload] of original) if (name !== "deck/z.xml") assert.equal(digest(payload), digest(current.get(name)), name);
    assert.ok(current.get("deck/z.xml").toString().includes('<x:opaque xmlns:x="urn:future" ref="2"> A &amp; B </x:opaque>'));
    assert.deepEqual(compared.differences.map(d => [d.code, d.part]), [["payload-changed", "/deck/z.xml"]]);
    const graph = await inspectOpcClosure(saved.bytes, { roots: ["/"] });
    assert.deepEqual(graph.relationships, closure.relationships); assert.deepEqual(graph.unreachedParts, closure.unreachedParts);
    if (round > 1) assert.equal(saved.bytes.length, previousSize);
    previousSize = saved.bytes.length; bytes = saved.bytes; session.dispose();
  }
});
