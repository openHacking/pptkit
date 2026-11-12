import assert from "node:assert/strict";
import test from "node:test";
import { importPptx, PptxImportError } from "../dist/index.js";
import { normalizePresentation, validatePresentation } from "@pptkit/core";
import { encodeStructure } from "./structure-fixture.mjs";
import { contentFixture, bytes, shape, body, transform } from "./content-fixture.mjs";

const open = (input = bytes(), extra = {}) => importPptx(input, { projection: "content", ...extra });
const errorCode = code => e => e instanceof PptxImportError && e.code === code;
const elements = session => normalizePresentation(session.document).slides[0].elements;
const unhandled = session => session.report.findings.filter(f => f.feature === "drawing.unmapped-property" && f.operation === "import");

test("explicit content projection imports direct rich text, frame, visibility and source identities in both dialects", async () => {
  for (const strict of [false, true]) {
    const session = await open(bytes(undefined, strict)), [text] = elements(session);
    assert.equal(text.type, "text"); assert.equal(text.plainText, " A & B ");
    assert.deepEqual(text.box, { x: 10, y: 20, width: 100, height: 50 });
    assert.equal(text.content[0].runs[0].style.fontFamily, "Test Font");
    assert.equal(text.content[0].runs[0].style.fontSize, 24);
    assert.equal(text.content[0].runs[0].style.bold, true);
    assert.equal(text.content[0].runs[0].style.color, "112233");
    assert.equal(text.content[0].runs[0].style.language, "zh-CN");
    assert.deepEqual(text.frame.margin, { left: 1, right: 2, top: 3, bottom: 4 });
    assert.equal(text.frame.verticalAlign, "middle"); assert.equal(text.frame.wrap, false);
    assert.equal(session.report.projection, "content");
    assert.equal(session.document.slides[0].hidden, true);
    assert.deepEqual(validatePresentation(session.document), []);
    const identity = session.report.identities.find(i => i.kind === "element");
    assert.equal(identity.id, "source-slide-300-object-2"); assert.equal(identity.source.sourceObjectId, "2"); assert.equal(identity.source.part, "/deck/z.xml");
    assert.match(identity.source.path, /spTree\[1\]\/p:sp\[1\]$/);
    assert.ok(Object.isFrozen(identity.source)); session.dispose();
  }
});

test("paragraph defaults, empty paragraphs, breaks, runs and numbered lists retain order and lexical text", async () => {
  const text = body(`<a:p><a:pPr algn="r" marL="25400" indent="-12700"><a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:spcBef><a:spcPts val="250"/></a:spcBef><a:spcAft><a:spcPts val="300"/></a:spcAft><a:buAutoNum type="alphaUcPeriod" startAt="3"/><a:defRPr sz="1800" b="1"/></a:pPr><a:r><a:t>α😀</a:t></a:r><a:br/><a:r><a:rPr b="0" i="1" u="sng" strike="sngStrike"/><a:t><![CDATA[e\u0301<&]]></a:t></a:r></a:p><a:p><a:endParaRPr sz="1200"/></a:p>`);
  const s = await open(bytes(shape({ txBox: true, text }))), [element] = elements(s), paragraphs = element.content;
  assert.equal(element.plainText, "α😀\ne\u0301<&\n");
  assert.equal(paragraphs.length, 2); assert.deepEqual(paragraphs[0].style, { align: "right", indent: 2, hanging: 1, lineSpacing: 1.5, spaceBefore: 2.5, spaceAfter: 3, bullet: { type: "number", style: "alphaUpperPeriod", startAt: 3 } });
  assert.equal(paragraphs[0].runs[0].style.bold, true);
  assert.equal(paragraphs[0].runs[2].style.bold, false); assert.equal(paragraphs[0].runs[2].style.italic, true);
  assert.equal(paragraphs[0].runs[2].style.underline, true); assert.equal(paragraphs[0].runs[2].style.strike, true);
  assert.equal(paragraphs[1].runs[0].style.fontSize, 12); s.dispose();
});

test("seven preset shapes, fill alpha, stroke, arrows and transforms become editable fields", async () => {
  const kinds = ["rect", "roundRect", "ellipse", "triangle", "diamond", "rightArrow", "chevron"];
  const markup = kinds.map((preset, i) => shape({ preset, id: String(i + 2), xfrm: transform('rot="2700000" flipH="1" flipV="true"'), fill: '<a:solidFill><a:srgbClr val="abcdef"><a:alpha val="50000"/></a:srgbClr></a:solidFill>', props: '<a:ln w="25400"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:prstDash val="dashDot"/><a:headEnd type="oval"/><a:tailEnd type="triangle"/></a:ln>' })).join("");
  const s = await open(bytes(markup));
  assert.deepEqual(elements(s).map(e => e.shape), ["rect", "roundRect", "ellipse", "triangle", "diamond", "arrow", "chevron"]);
  for (const e of elements(s)) {
    assert.deepEqual(e.style.fill, { type: "solid", color: "ABCDEF", opacity: .5 });
    assert.deepEqual(e.transform, { rotation: 45, flipH: true, flipV: true });
    assert.equal(e.style.stroke.width, 2); assert.equal(e.style.stroke.dash, "dashDot");
    assert.equal(e.style.stroke.beginArrow, "oval"); assert.equal(e.style.stroke.endArrow, "triangle");
  }
  s.dispose();
});

test("painted text frames remain shapes with editable rich text rather than dropping their fill", async () => {
  const s = await open(bytes(shape({ txBox: true, text: body(), fill: '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' })));
  const [e] = elements(s); assert.equal(e.type, "shape"); assert.equal(e.text.plainText, " A & B "); assert.equal(e.style.fill.color, "FF0000"); s.dispose();
});

function group(children, id = 3, origin = 5) {
  return `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="5400000">${'<a:off x="1270000" y="2540000"/><a:ext cx="2540000" cy="1270000"/>'}<a:chOff x="${origin * 12700}" y="${origin * 12700}"/><a:chExt cx="1270000" cy="635000"/></a:xfrm></p:grpSpPr>${children}</p:grpSp>`;
}
test("nested groups rebase child origins once while retaining scale and local rotations", async () => {
  const s = await open(bytes(group(group(shape(), 4, 2))));
  const [outer] = elements(s), inner = outer.children[0], child = inner.children[0];
  assert.deepEqual(outer.box, { x: 100, y: 200, width: 200, height: 100 });
  assert.deepEqual(outer.coordinateSize, { width: 100, height: 50 });
  assert.equal(inner.box.x, 95); assert.equal(inner.box.y, 195);
  assert.equal(child.box.x, 8); assert.equal(child.box.y, 18); assert.equal(child.box.width, 100);
  assert.equal(outer.transform.rotation, 90);
  assert.equal(s.report.identities.filter(i => i.kind === "element").length, 3); s.dispose();
});

test("unknown effects, properties, namespaces and siblings are located per object and per operation", async () => {
  const f = contentFixture(shape({ props: '<a:effectLst><a:outerShdw blurRad="100"/></a:effectLst>', nv: 'future="value"' }) + '<p:pic><p:nvPicPr><p:cNvPr id="3" name="Image"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr></p:pic>' + '<x:sp xmlns:x="urn:unknown"/>');
  const input = encodeStructure(f), s = await open(input);
  assert.equal(elements(s).length, 1);
  const fields = unhandled(s);
  assert.ok(fields.some(f => f.source.path.endsWith("/@future") && f.source.sourceObjectId === "2" && f.elementId === "source-slide-300-object-2"));
  assert.ok(fields.some(f => f.source.path.includes("effectLst")));
  assert.ok(fields.some(f => f.source.path.includes("x:sp")));
  const object = s.report.findings.filter(f => f.feature === "drawing.pic");
  assert.equal(object.length, 7); assert.ok(object.every(f => f.tier === (f.operation === "no-op-save" ? "preserved" : "unsupported")));
  const prepared = await s.prepareSave({ mode: "no-op" }); assert.equal(prepared.status, "ready");
  assert.equal(Buffer.compare(Buffer.from((await s.save(prepared.plan)).bytes), input), 0);
  assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked");
  assert.equal((await s.prepareSave({ mode: "rebuild" })).status, "blocked"); s.dispose();
});

test("supported text alpha/transforms map while unknown transforms and unresolved script/theme fonts remain reported", async () => {
  const text = body('<a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="112233"><a:alpha val="50000"/><a:tint val="10000"/></a:srgbClr></a:solidFill><a:latin typeface="+mn-lt"/></a:rPr><a:t>Theme</a:t></a:r><a:r><a:rPr><a:latin typeface="Latin"/><a:ea typeface="Other"/></a:rPr><a:t>東</a:t></a:r><a:r><a:rPr><a:solidFill><a:srgbClr val="112233"><a:hue val="120000"/></a:srgbClr></a:solidFill></a:rPr><a:t>Unknown transform</a:t></a:r></a:p>');
  const s = await open(bytes(shape({ txBox: true, text }))), fields = unhandled(s);
  const run = elements(s)[0].content[0].runs[0];
  assert.equal(run.style.color, "F4F4F4");
  assert.equal(run.style.colorOpacity, .5);
  assert.ok(fields.some(f => /Theme font|script-specific font/i.test(f.reason)));
  assert.ok(fields.some(f => /Unsupported color transform hue/.test(f.reason)));
  assert.ok(fields.some(f => /Script-specific/.test(f.reason)));
  assert.ok(fields.some(f => f.source.path.includes("hue")));
  assert.equal(fields.some(f => /tint|alpha/.test(f.source.path)), false);
  s.dispose();
});

test("missing local geometry or unsupported presets are omitted explicitly and source remains byte-preserved", async () => {
  const input = bytes(shape({ xfrm: "", id: "2" }) + shape({ preset: "cloud", id: "3" })), s = await open(input);
  assert.equal(elements(s).length, 0);
  assert.equal(s.report.findings.filter(f => f.feature === "drawing.sp" && f.operation === "import").length, 2);
  const p = await s.prepareSave({ mode: "no-op" }); assert.equal(Buffer.compare(Buffer.from((await s.save(p.plan)).bytes), input), 0); s.dispose();
});

test("drawing IDs are part-local but duplicate numeric identities including the root are rejected", async () => {
  for (const content of [shape({ id: "1" }), shape({ id: "2" }) + shape({ id: "02" }), group(shape({ id: "3" }))]) await assert.rejects(open(bytes(content)), errorCode("import-invalid-package"));
  const f = contentFixture(); f.parts["deck/a.xml"] = f.parts["deck/z.xml"];
  const s = await open(encodeStructure(f)); assert.notEqual(s.document.slides[0].elements[0].id, s.document.slides[1].elements[0].id); s.dispose();
});

test("malformed mapped numbers, booleans, text structure and conflicting choices fail before publishing a session", async () => {
  for (const content of [
    shape({ xfrm: transform('flipH="yes"') }),
    shape({ xfrm: transform().replace('cx="1270000"', 'cx="-1"') }),
    shape({ fill: '<a:noFill/><a:solidFill><a:srgbClr val="123456"/></a:solidFill>' }),
    shape({ text: body('<a:p><a:r/></a:p>') }),
    shape({ text: body('<a:p><a:pPr><a:buNone/><a:buChar char="x"/></a:pPr></a:p>') }),
    shape({ fill: '<a:solidFill><a:srgbClr val="invalid"/></a:solidFill>' }),
  ]) await assert.rejects(open(bytes(content)), errorCode("import-invalid-package"));
});

test("content import retains snapshot isolation, stale-plan rejection and exact restoration behavior", async () => {
  const source = bytes(), original = new Uint8Array(source), options = { projection: "content" };
  const pending = importPptx(source, options); source.fill(0); options.projection = "structure";
  const s = await pending, before = normalizePresentation(s.document);
  assert.equal(elements(s).length, 1);
  const prepared = await s.prepareSave({ mode: "no-op" }), slide = s.document.slides[0], element = slide.removeElement(slide.elements[0].id);
  await assert.rejects(s.save(prepared.plan), errorCode("save-snapshot-stale"));
  slide.addElement(element); assert.deepEqual(normalizePresentation(s.document), before);
  assert.deepEqual((await s.save(prepared.plan)).bytes, original); s.dispose();
});

test("budgets and cancellation cover content reports as well as archive and XML work", async () => {
  const input = bytes(Array.from({ length: 30 }, (_, i) => shape({ id: String(i + 2) })).join(""));
  await assert.rejects(open(input, { sessionLimits: { maxReportFindings: 220 } }), errorCode("import-limit-exceeded"));
  await assert.rejects(open(input, { xmlLimits: { maxNodes: 10 } }), errorCode("import-limit-exceeded"));
  await assert.rejects(open(input, { signal: AbortSignal.abort() }), errorCode("operation-aborted"));
  const controller = new AbortController(), pending = open(input, { signal: controller.signal }); controller.abort();
  await assert.rejects(pending, errorCode("operation-aborted"));
});

test("structure mode retains empty content and content sessions remain independent under concurrent import", async () => {
  const input = bytes(), structure = await importPptx(input, { projection: "structure" });
  assert.equal(structure.document.slides[0].elements.length, 0); structure.dispose();
  const [a, b] = await Promise.all([open(input), open(input)]); a.document.slides[0].removeElement(a.document.slides[0].elements[0].id);
  assert.equal(b.document.slides[0].elements.length, 1); assert.equal(b.report.identities.filter(i => i.kind === "element").length, 1);
  assert.ok(Object.isFrozen(b.report.findings)); a.dispose(); b.dispose();
});

test("MCE uses the effective branch and namespace aliases rather than XML prefix spelling", async () => {
  const f = contentFixture();
  f.parts["deck/z.xml"] = f.parts["deck/z.xml"]
    .replace(shape({ txBox: true, text: body() }), `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:q="urn:future"><mc:Choice Requires="q">${shape({ id: "3" })}</mc:Choice><mc:Fallback>${shape({ txBox: true, text: body() })}</mc:Fallback></mc:AlternateContent>`)
    .replaceAll("xmlns:p=", "xmlns:s=").replaceAll("<p:", "<s:").replaceAll("</p:", "</s:");
  const input = encodeStructure(f), s = await open(input);
  assert.equal(elements(s)[0].type, "text");
  assert.equal(elements(s)[0].plainText, " A & B ");
  assert.ok(s.report.identities.find(i => i.kind === "element").source.path.includes("/s:sp[1]"));
  assert.ok(s.report.inspection.markupCompatibility.some(entry => entry.part === "/deck/z.xml"));
  const saved = await s.save((await s.prepareSave({ mode: "no-op" })).plan);
  assert.equal(Buffer.compare(Buffer.from(saved.bytes), input), 0); s.dispose();
});

test("valid but unrepresentable margins, indent and group coordinate spaces report limitations", async () => {
  const text = body('<a:p><a:pPr indent="12700"/><a:r><a:t>Text</a:t></a:r></a:p>').replace('lIns="12700"', 'lIns="-12700"');
  const zeroGroup = group(shape({ id: "4" })).replace('<a:chExt cx="1270000"', '<a:chExt cx="0"');
  const s = await open(bytes(shape({ txBox: true, text }) + zeroGroup));
  assert.equal(elements(s).length, 1);
  assert.ok(unhandled(s).some(f => /Negative text insets/.test(f.reason)));
  assert.ok(unhandled(s).some(f => /Positive first-line/.test(f.reason)));
  assert.ok(s.report.findings.some(f => /Zero group child extent/.test(f.reason))); s.dispose();
});

test("duplicate IDs in an omitted group still reject, while opaque extension identities stay uninterpreted", async () => {
  const missingGeometry = group(shape({ id: "3" })).replace('<a:chExt cx="1270000"', '<a:chExt cx="0"');
  await assert.rejects(open(bytes(missingGeometry)), errorCode("import-invalid-package"));
  const s = await open(bytes(shape({ props: '<a:extLst><a:ext uri="urn:opaque"><p:sp><p:nvSpPr><p:cNvPr id="2"/></p:nvSpPr></p:sp></a:ext></a:extLst>' })));
  assert.equal(elements(s).length, 1); assert.ok(unhandled(s).some(f => f.source.path.includes("extLst"))); s.dispose();
});

test("projection getters are captured once and unsupported links remain located without fetching", async () => {
  let reads = 0;
  const f = contentFixture(shape({ nv: 'title="future"', props: 'non-schema character content', text: body('<a:p><a:r><a:rPr><a:hlinkClick r:id="missing"/></a:rPr><a:t>Link text</a:t></a:r></a:p>') }));
  const s = await importPptx(encodeStructure(f), { get projection() { reads++; if (reads > 1) throw new Error("Repeated projection read"); return "content"; } });
  assert.equal(reads, 1);
  assert.ok(unhandled(s).some(f => f.source.path.includes("hlinkClick")));
  assert.ok(unhandled(s).some(f => f.source.path.endsWith("text()")));
  assert.ok(unhandled(s).some(f => f.source.path.endsWith("/@title")));
  assert.equal(elements(s)[0].text.plainText, "Link text"); s.dispose();
});
