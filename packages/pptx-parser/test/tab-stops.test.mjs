import assert from "node:assert/strict";
import test from "node:test";
import { importPptx } from "../dist/index.js";
import { normalizePresentation } from "../../core/dist/index.js";
import { encodeStructure } from "./structure-fixture.mjs";
import { inheritedShape } from "./placeholder-fixture.mjs";
import { textBody, textStyleFixture } from "./text-style-fixture.mjs";

const xml = '<a:tabLst><a:tab pos="0" algn="l"/><a:tab pos="457200" algn="ctr"/><a:tab pos="914400" algn="r"/><a:tab pos="1828800" algn="dec"/></a:tabLst>';
const stops = [{ position: 0, align: "left" }, { position: 36, align: "center" }, { position: 72, align: "right" }, { position: 144, align: "decimal" }];
const open = (f, options = {}) => importPptx(encodeStructure(f), { projection: "content", ...options });
const body = p => textBody({ paragraphs: `<a:p><a:pPr>${p}</a:pPr><a:r><a:t>A\tB</a:t></a:r></a:p>` });
const local = (f, p) => f.set("deck/z.xml", "sld", inheritedShape({ text: body(p) }));
const style = s => normalizePresentation(s.document).slides[0].elements[0].text.content[0].style;
const row = s => s.report.textStyleInheritance.properties.find(p => p.property === "paragraph.tabStops" && p.target.part === "/deck/z.xml");

test("explicit tab lists project all alignments and source dependencies in both dialects", async () => {
  for (const strict of [false, true]) {
    const f = textStyleFixture(strict); local(f, xml); const s = await open(f);
    try {
      assert.deepEqual(style(s).tabStops, stops); assert.equal(row(s).level, "paragraph");
      assert.equal(row(s).status, "mapped"); assert.equal(row(s).resolvedFrom.length, 4);
      assert.match(row(s).source.path, /a:tabLst\[1\]$/); assert.ok(Object.isFrozen(row(s).value));
    } finally { s.dispose(); }
  }
});

test("inherited tab lists are replaced or cleared by a local list; unsupported local choices block inheritance", async () => {
  for (const [localTabs, expected, status] of [[undefined, stops, "mapped"], ["<a:tabLst/>", [], "mapped"], ['<a:tabLst><a:tab pos="100" algn="r"/></a:tabLst>', [{ position: 100 / 12700, align: "right" }], "mapped"], ['<a:tabLst><a:tab pos="100"/></a:tabLst>', undefined, "unsupported"], ['<a:tabLst><a:tab pos="100" algn="constructor"/></a:tabLst>', undefined, "unsupported"], ['<a:tabLst><a:tab pos="100" algn="l"/><a:tab pos="100" algn="r"/></a:tabLst>', undefined, "unsupported"]]) {
    const f = textStyleFixture();
    f.set("deck/layout.xml", "sldLayout", inheritedShape({ text: textBody({ list: `<a:lvl1pPr>${xml}</a:lvl1pPr>` }) }));
    local(f, localTabs ?? ""); const s = await open(f);
    try { assert.deepEqual(style(s).tabStops, expected); assert.equal(row(s).status, status); assert.equal(row(s).level, localTabs === undefined ? "layout-placeholder" : "paragraph"); }
    finally { s.dispose(); }
  }
});

test("invalid numeric tab positions reject and oversized lists block without partial projection", async () => {
  const bad = textStyleFixture(); local(bad, '<a:tabLst><a:tab pos="not-an-integer" algn="l"/></a:tabLst>');
  await assert.rejects(open(bad), e => e.code === "import-invalid-package");
  const oversized = textStyleFixture(); local(oversized, `<a:tabLst>${Array.from({ length: 257 }, (_, i) => `<a:tab pos="${i}" algn="l"/>`).join("")}</a:tabLst>`);
  const s = await open(oversized); try { assert.equal(row(s).status, "unsupported"); assert.equal(style(s).tabStops, undefined); } finally { s.dispose(); }
});

test("tab-stop provenance charges the cumulative dependency budget and no-op save preserves bytes", async () => {
  const f = textStyleFixture(); local(f, xml); const bytes = encodeStructure(f), s = await importPptx(bytes, { projection: "content" });
  const count = s.report.textStyleInheritance.properties.reduce((sum, p) => sum + (p.resolvedFrom?.length ?? 0), 0);
  assert.ok(count >= 4);
  const plan = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(plan.plan)).bytes, new Uint8Array(bytes)); s.dispose();
  await assert.rejects(open(f, { sessionLimits: { maxTextStyleDependencies: count - 1 } }), { code: "import-limit-exceeded" });
  const exact = await open(f, { sessionLimits: { maxTextStyleDependencies: count } }); exact.dispose();
});
