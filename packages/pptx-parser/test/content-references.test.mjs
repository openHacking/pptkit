import assert from "node:assert/strict";
import test from "node:test";
import { normalizePresentation, validatePresentation } from "@pptkit/core";
import { importPptx } from "../dist/index.js";
import { encodeStructure } from "./structure-fixture.mjs";
import { shape, body, transform } from "./content-fixture.mjs";
import { connector, group, click, withClick, referenceFixture } from "./reference-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";
import { imageFixture, picture, encodeImages } from "./image-fixture.mjs";

const open = (content, strict = false, extras = [], options = {}) => importPptx(encodeStructure(referenceFixture(content, strict, extras)), { projection: "content", ...options });
const elements = s => normalizePresentation(s.document).slides[0].elements;
const reports = s => s.report.findings.filter(f => f.operation === "import");
const closePoint = (point, expected) => { assert.ok(Math.abs(point.x - expected.x) < 1e-9, `${point.x} vs ${expected.x}`); assert.ok(Math.abs(point.y - expected.y) < 1e-9, `${point.y} vs ${expected.y}`); };

test("native straight connectors bind forward/backward source IDs and preserve direct stroke fields in both dialects", async () => {
  for (const strict of [false, true]) {
    const s = await open(connector({ startId: 3, endId: 4, endSite: 2 }) + shape({ id: "3" }) + shape({ id: "4", preset: "ellipse" }) + connector({ id: 5, startId: 4, startSite: 6, endId: 3, endSite: 1 }), strict);
    const [a, rect, ellipse, b] = elements(s);
    assert.equal(a.type, "connector");
    assert.deepEqual(a.start, { elementId: rect.id, anchor: "right" }); assert.deepEqual(a.end, { elementId: ellipse.id, anchor: "left" });
    // Ellipse uses 0/2/4/6 rather than rectangle's 0/1/2/3.
    assert.deepEqual(b.start, { elementId: ellipse.id, anchor: "right" }); assert.deepEqual(b.end, { elementId: rect.id, anchor: "left" });
    assert.deepEqual(a.style.paint, { type: "solid", color: "123456", opacity: 1 });
    assert.equal(a.style.width, 2); assert.equal(a.style.dash, "dashDot"); assert.equal(a.style.beginArrow, "oval"); assert.equal(a.style.endArrow, "triangle");
    assert.deepEqual(a.transform, { rotation: 0, flipH: false, flipV: false });
    assert.deepEqual(validatePresentation(s.document), []);
    assert.equal(s.report.diagnostics.filter(d => d.code === "connector-attachment-unmapped").length, 0);
    assert.ok(s.report.identities.some(i => i.id === a.id && i.source.sourceObjectId === "2" && i.source.path.endsWith("/p:cxnSp[1]")));
    assert.equal(reports(s).filter(f => f.feature === "drawing.connector.attachment").length, 4); s.dispose();
  }
});

test("every supported native site maps independently and unsupported indices never wrap around", async () => {
  for (const [preset, indices] of [["rect", [0, 1, 2, 3]], ["roundRect", [0, 1, 2, 3]], ["ellipse", [0, 2, 4, 6]]]) {
    const content = indices.map((site, i) => connector({ id: i + 2, startId: 20, startSite: site })).join("") + shape({ id: "20", preset });
    const s = await open(content);
    assert.deepEqual(elements(s).slice(0, 4).map(e => e.start.anchor), ["top", "left", "bottom", "right"]); s.dispose();
  }
  const s = await open(connector({ startId: 3, startSite: 4294967295 }) + shape({ id: "3" }));
  assert.deepEqual(elements(s)[0].start, { x: 10, y: 20 });
  assert.match(s.report.diagnostics.find(d => d.code === "connector-attachment-unmapped").message, /site 4294967295/); s.dispose();
});

test("free straight endpoints flatten source rotation/flips once, including zero extents and reversed directions", async () => {
  const cases = [
    ["", { x: 10, y: 20 }, { x: 110, y: 70 }],
    ['flipH="1"', { x: 110, y: 20 }, { x: 10, y: 70 }],
    ['flipV="1"', { x: 10, y: 70 }, { x: 110, y: 20 }],
    ['rot="5400000"', { x: 85, y: -5 }, { x: 35, y: 95 }],
    ['rot="5400000" flipH="1" flipV="1"', { x: 35, y: 95 }, { x: 85, y: -5 }],
  ];
  for (const [extra, start, end] of cases) {
    const s = await open(connector({ preset: "line", xfrm: transform(extra) })), [e] = elements(s);
    closePoint(e.start, start); closePoint(e.end, end);
    assert.equal(e.transform.rotation, 0); assert.equal(e.transform.flipH, false); assert.equal(e.transform.flipV, false); s.dispose();
  }
  const zero = await open(connector({ xfrm: transform().replace('cx="1270000"', 'cx="0"').replace('cy="635000"', 'cy="0"') }));
  assert.deepEqual(elements(zero)[0].start, elements(zero)[0].end); assert.deepEqual(validatePresentation(zero.document), []); zero.dispose();
});

test("bentConnector2 imports its transformed elbow and same-scope attachments in both dialects", async () => {
  const cases = [
    ['', { x: 10, y: 20 }, { x: 110, y: 20 }, { x: 110, y: 70 }],
    ['flipH="1"', { x: 110, y: 20 }, { x: 10, y: 20 }, { x: 10, y: 70 }],
    ['flipV="1"', { x: 10, y: 70 }, { x: 110, y: 70 }, { x: 110, y: 20 }],
    ['flipH="1" flipV="1"', { x: 110, y: 70 }, { x: 10, y: 70 }, { x: 10, y: 20 }],
    ['rot="5400000"', { x: 85, y: -5 }, { x: 85, y: 95 }, { x: 35, y: 95 }],
  ];
  for (const strict of [false, true]) for (const [extra, start, bend, end] of cases) {
    const s = await open(connector({ preset: 'bentConnector2', xfrm: transform(extra) }), strict), [item] = elements(s);
    assert.equal(item.type, 'connector'); closePoint(item.start, start); closePoint(item.route[0], bend); closePoint(item.end, end);
    assert.deepEqual(item.transform, { rotation: 0, flipH: false, flipV: false }); assert.deepEqual(validatePresentation(s.document), []);
    assert.ok(!reports(s).some(f => f.feature === 'drawing.cxnSp' && f.tier === 'unsupported')); s.dispose();
  }
  const attached = await open(connector({ preset: 'bentConnector2', startId: 3, endId: 4 }) + shape({ id: '3' }) + shape({ id: '4' }));
  const [line, first, second] = elements(attached); assert.deepEqual(line.start, { elementId: first.id, anchor: 'right' }); assert.deepEqual(line.end, { elementId: second.id, anchor: 'left' });
  assert.deepEqual(line.route, [{ x: 110, y: 20 }]); attached.dispose();
  const modified = connector({ preset: 'bentConnector2' }).replace('<a:avLst/>', '<a:avLst><a:gd name="adj1" fmla="val 50000"/></a:avLst>');
  const unsupported = await open(modified); assert.deepEqual(elements(unsupported), []); assert.ok(reports(unsupported).some(f => f.feature === 'drawing.cxnSp' && f.tier === 'unsupported')); unsupported.dispose();
  const nested = await open(group(group(connector({ preset: 'bentConnector2', startId: 3 }) + shape({ id: '3' }), 11, 2), 10, 5));
  const nestedGroup = elements(nested)[0].children[0], [nestedLine, nestedTarget] = nestedGroup.children;
  assert.deepEqual(nestedLine.start, { elementId: nestedTarget.id, anchor: 'right' }); assert.deepEqual(nestedLine.route, [{ x: 108, y: 18 }]); assert.deepEqual(nestedLine.end, { x: 108, y: 68 });
  assert.deepEqual(validatePresentation(nested.document), []); nested.dispose();
  const diagonal = await open(connector({ preset: 'bentConnector2', xfrm: transform('rot="2700000"') })), [turned] = elements(diagonal);
  const root = Math.SQRT1_2;
  closePoint(turned.start, { x: 60 - 25 * root, y: 45 - 75 * root });
  closePoint(turned.route[0], { x: 60 + 75 * root, y: 45 + 25 * root });
  closePoint(turned.end, { x: 60 + 25 * root, y: 45 + 75 * root });
  assert.deepEqual(validatePresentation(diagonal.document), []); diagonal.dispose();
});

test('bentConnector3 maps default and bounded literal adjustments, preserving transforms and rejecting ambiguous guides', async () => {
  for (const strict of [false, true]) for (const [value, expected] of [[undefined, 60], [0, 10], [25000, 35], [100000, 110]]) {
    const base = connector({ preset: 'bentConnector3' });
    const source = value === undefined ? base : base.replace('<a:avLst/>', `<a:avLst><a:gd name="adj1" fmla="val ${value}"/></a:avLst>`);
    const s = await open(source, strict), [item] = elements(s);
    assert.deepEqual(item.route, [{ x: expected, y: 20 }, { x: expected, y: 70 }]); assert.deepEqual(validatePresentation(s.document), []); s.dispose();
  }
  const flipped = await open(connector({ preset: 'bentConnector3', xfrm: transform('flipH="1" flipV="1"') }).replace('<a:avLst/>', '<a:avLst><a:gd name="adj1" fmla="val 25000"/></a:avLst>'));
  assert.deepEqual(elements(flipped)[0].route, [{ x: 85, y: 70 }, { x: 85, y: 20 }]); flipped.dispose();
  for (const content of ['<a:gd name="adj1" fmla="val 100001"/>', '<a:gd name="adj2" fmla="val 50000"/>', '<a:gd name="adj1" fmla="*/ w 1 2"/>', '<a:gd name="adj1" fmla="val 1"/><a:gd name="adj1" fmla="val 2"/>', '<a:gd name="adj1" fmla="val 1" extra="1"/>']) {
    const s = await open(connector({ preset: 'bentConnector3' }).replace('<a:avLst/>', `<a:avLst>${content}</a:avLst>`));
    assert.deepEqual(elements(s), []); assert.ok(reports(s).some(item => item.feature === 'drawing.cxnSp' && item.tier === 'unsupported')); s.dispose();
  }
});

test("nested group endpoints are rebased once and attach within their immediate local scope", async () => {
  const s = await open(group(group(connector({ startId: 3 }) + shape({ id: "3" }), 11, 2), 10, 5));
  const e = elements(s)[0].children[0], [line, target] = e.children;
  assert.deepEqual(line.start, { elementId: target.id, anchor: "right" });
  assert.deepEqual(line.end, { x: 108, y: 68 });
  assert.equal(target.box.x, 8); assert.equal(target.box.y, 18);
  assert.deepEqual(validatePresentation(s.document), []); s.dispose();
});

test("missing/omitted/cross-group/diagonal/non-preset attachments retain free points with located fallback diagnostics", async () => {
  for (const content of [
    connector({ startId: 99 }),
    connector({ startId: 3 }) + shape({ id: "3", preset: "cloud" }),
    connector({ startId: 3 }) + group(shape({ id: "3" })),
    connector({ startId: 3, startSite: 1 }) + shape({ id: "3", preset: "ellipse" }),
    connector({ startId: 3 }) + shape({ id: "3", preset: "triangle" }),
    connector({ startId: 2 }),
  ]) {
    const s = await open(content), [e] = elements(s);
    assert.deepEqual(e.start, { x: 10, y: 20 });
    assert.deepEqual(validatePresentation(s.document), []);
    const warning = s.report.diagnostics.find(d => d.code === "connector-attachment-unmapped");
    assert.equal(warning.elementId, e.id); assert.equal(warning.source.sourceObjectId, "2");
    assert.ok(warning.source.path.endsWith("a:stCxn[1]"));
    assert.ok(reports(s).some(f => f.feature === "drawing.connector.attachment" && f.tier === "fallback")); s.dispose();
  }
});

test("unknown routes remain omitted and connector property extensions stay independently reported", async () => {
  const s = await open(connector({ preset: "bentConnector4" }) + connector({ id: 5, preset: "curvedConnector3" }) + connector({ id: 6, nv: '<a:cxnSpLocks noMove="1"/>', props: '<a:effectLst/>' }) + shape({ id: "3" }));
  assert.deepEqual(elements(s).map(e => e.type), ["connector", "shape"]);
  assert.equal(reports(s).filter(f => f.feature === "drawing.cxnSp" && f.tier === "unsupported").length, 2);
  for (const name of ["cxnSpLocks", "effectLst"]) assert.ok(reports(s).some(f => f.source.path?.includes(name) && f.elementId === elements(s)[0].id)); s.dispose();
});

test("invalid connection IDs/sites and duplicate endpoints reject before a document is returned", async () => {
  for (const content of [
    connector({ startId: "-1" }), connector({ startId: "4294967296" }), connector({ startId: "bad" }), connector({ startId: 3, startSite: "-1" }),
    connector({ startId: 3, startSite: "1.5" }), connector({ startId: 3, nv: '<a:stCxn id="3" idx="0"/>' }),
    connector({ startId: 3 }).replace(' idx="3"', ""),
    connector() + shape({ id: "2" }),
  ]) await assert.rejects(open(content), { code: "import-invalid-package" });
});

test("object and run URL/slide actions import exact targets, tooltips and source relation provenance", async () => {
  for (const strict of [false, true]) {
    const text = body(`<a:p><a:r><a:rPr>${click("jump", 'action="ppaction://hlinksldjump" tooltip="Go &amp; see"')}</a:rPr><a:t>Next</a:t></a:r><a:r><a:rPr>${click("url")}</a:rPr><a:t>URL</a:t></a:r></a:p>`);
    const s = await open(withClick(shape({ txBox: true, text }), click("self", 'tooltip="Self"')), strict), [e] = elements(s);
    assert.deepEqual(e.action, { type: "slide", slideId: s.document.slides[0].id, tooltip: "Self" });
    assert.deepEqual(e.content[0].runs[0].action, { type: "slide", slideId: s.document.slides[1].id, tooltip: "Go & see" });
    assert.deepEqual(e.content[0].runs[1].action, { type: "url", url: "https://example.invalid/a?x=1&y=%E4%B8%AD#part" });
    assert.equal(e.plainText, "NextURL"); assert.deepEqual(validatePresentation(s.document), []);
    const source = s.report.findings.find(f => f.feature === "drawing.click-action" && f.operation === "import" && f.source.relationshipId === "jump");
    assert.equal(source.elementId, e.id); assert.ok(source.source.path.includes("a:rPr[1]/a:hlinkClick[1]")); assert.equal(source.tier, "editable");
    assert.equal(s.report.findings.find(f => f.feature === "drawing.click-action" && f.operation === "preview" && f.source.relationshipId === "jump").tier, "fallback");
    assert.equal(s.report.identities.filter(i => i.kind === "element").length, 1); s.dispose();
  }
});

test("click actions on groups, shapes and native connectors map without inheriting a parent action", async () => {
  const s = await open(withClick(group(withClick(shape({ id: "3" }), click("jump"))), click("self")) + withClick(connector({ id: 5 }), click("url")));
  const [g, line] = elements(s); assert.equal(g.action.type, "slide"); assert.equal(g.children[0].action.slideId, s.document.slides[1].id);
  assert.equal(line.action.type, "url"); s.dispose();
});

test("unsupported commands, relative/unsafe links and unresolved actions are diagnostic-only and never become live actions", async () => {
  const samples = [
    { id: "missing" }, { id: "jump", attrs: 'action="ppaction://macro?name=Bad"' },
    { id: "url", attrs: 'action="ppaction://hlinkfile"' },
    { id: "extra", rel: f => ({ ...f.rel("extra", "hyperlink", "javascript:alert(1)"), TargetMode: "External" }) },
    { id: "extra", rel: f => ({ ...f.rel("extra", "hyperlink", "file:///tmp/unknown"), TargetMode: "External" }) },
    { id: "extra", rel: f => ({ ...f.rel("extra", "hyperlink", "../relative.html"), TargetMode: "External" }) },
    { id: "extra", rel: f => ({ ...f.rel("extra", "slide", "https://example.invalid/a.pptx"), TargetMode: "External" }) },
    { id: "extra", rel: f => f.rel("extra", "slide", "a.xml#shape") },
    { id: "extra", rel: f => f.rel("extra", "hyperlink", "a.xml") },
    { id: "extra", rel: f => f.rel("extra", "image", "a.xml") },
  ];
  for (const sample of samples) {
    const s = await open(withClick(shape(), click(sample.id, sample.attrs)), false, sample.rel ? [sample.rel] : []);
    assert.equal(elements(s)[0].action, undefined); assert.deepEqual(validatePresentation(s.document), []);
    const warning = s.report.diagnostics.find(d => d.code === "click-action-unmapped");
    assert.equal(warning.source.relationshipId, sample.id); assert.equal(warning.elementId, elements(s)[0].id);
    assert.ok(reports(s).some(f => f.feature === "drawing.click-action" && f.tier === "unsupported"));
    const prepared = await s.prepareSave({ mode: "no-op" }); assert.equal(prepared.status, "ready"); s.dispose();
  }
});

test("HTTP/HTTPS/mailto spellings remain exact and hover/sound/unknown attributes remain separately visible", async () => {
  for (const target of ["HTTPS://EXAMPLE.invalid/%41?q=%2F#frag", "mailto:user@example.invalid?subject=Hello%20world"]) {
    const link = click("extra", 'tooltip="" history="0"', '<a:snd r:embed="unread"/>') + '<a:hlinkHover r:id="jump"/>';
    const s = await open(withClick(shape(), link), false, [f => ({ ...f.rel("extra", "hyperlink", target), TargetMode: "External" })]);
    assert.deepEqual(elements(s)[0].action, { type: "url", url: target, tooltip: "" });
    for (const suffix of ["@history", "a:snd[1]", "a:hlinkHover[1]"]) assert.ok(reports(s).some(f => f.source.path?.endsWith(suffix)));
    s.dispose();
  }
});

test("slides outside the effective order are not fabricated as hyperlink targets", async () => {
  const f = referenceFixture(withClick(shape(), click("orphan")));
  f.parts["deck/orphan.xml"] = f.p("sld", "<p:cSld/>"); f.types["deck/orphan.xml"] = "presentationml.slide";
  f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), f.rel("orphan", "slide", "orphan.xml")]);
  const s = await importPptx(encodeStructure(f), { projection: "content" });
  assert.equal(s.document.slides.length, 2); assert.equal(elements(s)[0].action, undefined);
  assert.match(s.report.diagnostics.find(d => d.code === "click-action-unmapped").message, /outside the effective/); s.dispose();
});

test("namespace aliases and MCE-selected links/connectors use expanded names with stable identity mapping", async () => {
  const f = referenceFixture(connector({ startId: 3 }) + withClick(shape({ id: "3" }), click("jump")), true);
  f.parts["deck/z.xml"] = f.parts["deck/z.xml"]
    .replaceAll("xmlns:r=", "xmlns:rel=").replaceAll("r:id=", "rel:id=")
    .replace(click("jump").replace("r:id", "rel:id"), `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:future"><mc:Choice Requires="u">${click("missing").replace("r:id", "rel:id")}</mc:Choice><mc:Fallback>${click("jump").replace("r:id", "rel:id")}</mc:Fallback></mc:AlternateContent>`);
  const s = await importPptx(encodeStructure(f), { projection: "content" }), [line, target] = elements(s);
  assert.equal(line.start.elementId, target.id); assert.equal(target.action.slideId, s.document.slides[1].id);
  assert.equal(reports(s).filter(f => f.source.path?.endsWith("/@rel:id")).length, 0); s.dispose();
});

test("picture click actions share the mapped image identity without implicitly loading the link", async () => {
  const f = imageFixture(withClick(picture(), click("jump", 'action="ppaction://hlinksldjump"')));
  f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), f.rel("image", "image", "../media/raster.bin"), f.rel("jump", "slide", "a.xml")]);
  const s = await importPptx(encodeImages(f), { projection: "content" }), [image] = elements(s);
  assert.equal(image.type, "image"); assert.equal(image.action.slideId, s.document.slides[1].id);
  assert.equal(s.document.assets.length, 1); assert.deepEqual(validatePresentation(s.document), []); s.dispose();
});

test("insertion/default hyperlink properties are not misapplied to existing text; duplicate click choices reject", async () => {
  const text = body(`<a:p><a:pPr><a:defRPr>${click("url")}</a:defRPr></a:pPr><a:r><a:t>Plain</a:t></a:r><a:endParaRPr>${click("jump")}</a:endParaRPr></a:p>`);
  const s = await open(shape({ txBox: true, text })), [e] = elements(s);
  assert.equal(e.content[0].runs[0].action, undefined);
  assert.ok(reports(s).some(f => f.source.path?.includes("defRPr[1]/a:hlinkClick[1]")));
  assert.ok(reports(s).some(f => f.source.path?.includes("endParaRPr[1]/a:hlinkClick[1]"))); s.dispose();
  await assert.rejects(open(withClick(shape(), click("url") + click("jump"))), { code: "import-invalid-package" });
});

test("deleting an imported target leaves a diagnosable reference and blocks saving instead of retargeting", async () => {
  const s = await open(connector({ startId: 3 }) + shape({ id: "3" }));
  const slide = s.document.slides[0], target = slide.removeElement(slide.elements[1].id);
  assert.ok(validatePresentation(s.document).some(d => d.code === "missing-connector-target"));
  assert.equal(slide.elements[0].start.elementId, target.id);
  const blocked = await s.prepareSave({ mode: "no-op" }); assert.equal(blocked.status, "blocked");
  slide.addElement(target); assert.deepEqual(validatePresentation(s.document), []);
  assert.equal((await s.prepareSave({ mode: "no-op" })).status, "ready"); s.dispose();
});

test("copy/delete, source snapshots, stale plans, budgets and cancellation keep reference imports isolated", async () => {
  const f = referenceFixture(connector({ startId: 3 }) + withClick(shape({ id: "3" }), click("self"))), bytes = encodeStructure(f), original = new Uint8Array(bytes);
  const pending = importPptx(bytes, { projection: "content" }); bytes.fill(0);
  const s = await pending, source = s.document.slides[0], plan = await s.prepareSave({ mode: "no-op" });
  const copied = s.document.duplicateSlide(source.id); s.document.removeSlide(source.id);
  assert.equal(copied.elements[0].start.elementId, copied.elements[1].id); assert.equal(copied.elements[1].action.slideId, copied.id);
  assert.deepEqual(validatePresentation(s.document), []);
  await assert.rejects(s.save(plan.plan), { code: "save-snapshot-stale" });
  assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked"); s.dispose();
  await assert.rejects(importPptx(original, { projection: "content", sessionLimits: { maxReportFindings: 100 } }), { code: "import-limit-exceeded" });
  const controller = new AbortController(), cancelled = importPptx(original, { projection: "content", signal: controller.signal }); controller.abort();
  await assert.rejects(cancelled, { code: "operation-aborted" });
  const structure = await importPptx(original, { projection: "structure" }); assert.equal(structure.document.slides[0].elements.length, 0); structure.dispose();
});
