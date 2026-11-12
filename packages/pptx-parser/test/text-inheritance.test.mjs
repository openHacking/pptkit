import test from "node:test";
import assert from "node:assert/strict";
import { importPptx, DEFAULT_SESSION_LIMITS } from "../dist/index.js";
import { normalizePresentation, validatePresentation } from "../../core/dist/index.js";
import { encodeStructure } from "./structure-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";
import { shape } from "./content-fixture.mjs";
import { inheritedShape, placeholderTransform } from "./placeholder-fixture.mjs";
import { textBody, textStyleFixture, masterList, layoutList, setTextTheme } from "./text-style-fixture.mjs";
import { notesFixture, notesShape } from "./notes-fixture.mjs";
import { tableFixture, table, tableRow, cell } from "./table-fixture.mjs";

const open = (f, options = {}) => importPptx(encodeStructure(f), { projection: "content", ...options });
const projected = s => normalizePresentation(s.document).slides[0].elements[0].text;
const records = (s, key, part = "/deck/z.xml") => s.report.textStyleInheritance.properties.filter(p => p.property === key && p.target.part === part);
const code = expected => error => error.code === expected;
const setLocal = (f, value) => f.set("deck/z.xml", "sld", inheritedShape({ text: value }));

test("paragraph direction inherits and explicit LTR overrides RTL in both dialects", async () => {
  for (const strict of [false, true]) {
    const f = textStyleFixture(strict);
    f.parts['deck/layout.xml'] = f.parts['deck/layout.xml'].replace('<a:lvl1pPr>', '<a:lvl1pPr rtl="1">');
    setLocal(f, textBody({ paragraphs: '<a:p><a:r><a:t>שלום 123</a:t></a:r></a:p><a:p><a:pPr rtl="0"/><a:r><a:t>ABC</a:t></a:r></a:p>' }));
    const session = await open(f);
    try {
      assert.deepEqual(projected(session).content.map(p => p.style.direction), ['rtl', 'ltr']);
      const declarations = records(session, 'paragraph.direction');
      assert.ok(declarations.some(item => item.value === 'rtl' && item.source.part === '/deck/layout.xml'));
      assert.ok(declarations.some(item => item.value === 'ltr' && item.source.part === '/deck/z.xml'));
    } finally { session.dispose(); }
  }
});

test("presentation direction levels resolve before default LTR is elided", async () => {
  for (const strict of [false, true]) for (const [defaultRtl, levelRtl, expected] of [['1', '0', undefined], ['0', '1', 'rtl']]) {
    const f = textStyleFixture(strict);
    f.parts['deck/main.data'] = f.parts['deck/main.data'].replace('</p:presentation>', `<p:defaultTextStyle><a:defPPr rtl="${defaultRtl}"/><a:lvl1pPr rtl="${levelRtl}"/></p:defaultTextStyle></p:presentation>`).replace('xmlns:r=', `xmlns:a="${f.A}" xmlns:r=`);
    f.set('deck/z.xml', 'sld', shape({ text: textBody() }));
    const session = await open(f);
    try {
      assert.equal(projected(session).content[0].style.direction, expected);
      assert.ok(records(session, 'paragraph.direction').some(item => item.level === 'presentation-default' && item.value === (levelRtl === '1' ? 'rtl' : 'ltr')));
    } finally { session.dispose(); }
  }
});

test("paragraph direction accepts XML boolean spellings and rejects malformed values", async () => {
  for (const [raw, expected] of [['true', 'rtl'], ['1', 'rtl'], ['false', 'ltr'], ['0', 'ltr']]) {
    const f = textStyleFixture();
    setLocal(f, textBody({ paragraphs: `<a:p><a:pPr rtl="${raw}"/><a:r><a:t>ABC</a:t></a:r></a:p>` }));
    const session = await open(f);
    try { assert.equal(projected(session).content[0].style.direction, expected); }
    finally { session.dispose(); }
  }
  const f = textStyleFixture();
  setLocal(f, textBody({ paragraphs: '<a:p><a:pPr rtl="invalid"/><a:r><a:t>ABC</a:t></a:r></a:p>' }));
  await assert.rejects(open(f), code('import-invalid-package'));
});

test("run RTL inherits, accepts explicit false and validates its XML boolean", async () => {
  for (const strict of [false, true]) {
    const f = textStyleFixture(strict);
    f.parts['deck/layout.xml'] = f.parts['deck/layout.xml'].replace('</a:defRPr>', '<a:rtl/></a:defRPr>');
    setLocal(f, textBody({ paragraphs: '<a:p><a:r><a:t>مرحبا</a:t></a:r><a:r><a:rPr><a:rtl val="0"/></a:rPr><a:t>ABC</a:t></a:r></a:p>' }));
    const session = await open(f);
    try {
      assert.deepEqual(projected(session).content[0].runs.map(run => run.style.rightToLeft), [true, false]);
      const values = records(session, 'run.rightToLeft');
      assert.ok(values.some(item => item.value === true && item.source.part === '/deck/layout.xml'));
      assert.ok(values.some(item => item.value === false && item.source.part === '/deck/z.xml'));
    } finally { session.dispose(); }
  }
  for (const raw of ['1', 'true', '0', 'false']) {
    const f = textStyleFixture();
    setLocal(f, textBody({ paragraphs: `<a:p><a:r><a:rPr><a:rtl val="${raw}"/></a:rPr><a:t>ABC</a:t></a:r></a:p>` }));
    const session = await open(f);
    try { assert.equal(projected(session).content[0].runs[0].style.rightToLeft, raw === '1' || raw === 'true'); }
    finally { session.dispose(); }
  }
  const f = textStyleFixture();
  setLocal(f, textBody({ paragraphs: '<a:p><a:r><a:rPr><a:rtl val="invalid"/></a:rPr><a:t>ABC</a:t></a:r></a:p>' }));
  await assert.rejects(open(f), code('import-invalid-package'));
});

test("placeholder frame/paragraph/run styles cascade property by property in both source dialects", async () => {
  for (const strict of [false, true]) {
    const s = await open(textStyleFixture(strict)), text = projected(s), [paragraph] = text.content, [run] = paragraph.runs;
    assert.deepEqual(validatePresentation(s.document), []);
    assert.deepEqual(text.frame.margin, { left: 5, right: 2, top: 0, bottom: 4 });
    assert.equal(text.frame.verticalAlign, "bottom"); assert.equal(text.frame.wrap, true);
    assert.deepEqual(text.frame.autoFit, { mode: "shrink", fontScale: .75 });
    assert.equal(paragraph.style.align, "center"); assert.equal(paragraph.style.indent, 2);
    assert.equal(paragraph.style.hanging, 1); assert.equal(paragraph.style.spaceBefore, 1.5);
    assert.deepEqual(paragraph.style.bullet, { type: "bullet", character: "*" });
    assert.equal(run.style.fontSize, 20); assert.equal(run.style.fontFamily, "Parent Font");
    assert.equal(run.style.bold, true); assert.equal(run.style.italic, false);
    assert.equal(run.style.underline, true); assert.equal(run.style.strike, true);
    assert.equal(run.style.color, "445566"); assert.equal(run.style.language, "zh-CN");
    assert.equal(text.plainText, "Inherited α😀");
    assert.doesNotMatch(JSON.stringify(normalizePresentation(s.document)), /MASTER PROMPT/);
    assert.equal(records(s, "run.fontSize")[0].level, "layout-placeholder");
    assert.equal(records(s, "run.fontSize")[0].source.part, "/deck/layout.xml");
    assert.match(records(s, "run.fontSize")[0].source.path, /a:lstStyle\[1\]\/a:lvl1pPr\[1\]\/a:defRPr\[1\]\/@sz$/);
    assert.equal(records(s, "run.fontFamily")[0].source.sourceObjectId, "2");
    assert.equal(records(s, "frame.margin.top")[0].value, 0);
    assert.ok(Object.isFrozen(records(s, "run.color")[0].target)); s.dispose();
  }
});

test("local frame, list, paragraph and run declarations override only their own properties, including explicit clearing", async () => {
  const f = textStyleFixture();
  setLocal(f, textBody({
    frame: 'rIns="0" anchor="t" wrap="none"', fit: "<a:noAutofit/>",
    list: '<a:defPPr><a:defRPr b="0"/></a:defPPr><a:lvl1pPr><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buNone/></a:lvl1pPr>',
    paragraphs: '<a:p><a:pPr algn="l" marL="0" indent="0"><a:defRPr u="none" strike="noStrike"/></a:pPr><a:r><a:rPr sz="1100" b="0"/><a:t>First</a:t></a:r><a:r><a:t>Second</a:t></a:r></a:p>',
  }));
  const s = await open(f), text = projected(s), paragraph = text.content[0];
  assert.deepEqual(text.frame.margin, { left: 5, right: 0, top: 0, bottom: 4 });
  assert.equal(text.frame.wrap, false); assert.equal(text.frame.verticalAlign, "top"); assert.deepEqual(text.frame.autoFit, { mode: "none" });
  assert.equal(paragraph.style.align, "left"); assert.equal(paragraph.style.indent, 0); assert.equal(paragraph.style.hanging, 0); assert.equal(paragraph.style.spaceAfter, 0);
  assert.deepEqual(paragraph.style.bullet, { type: "none" });
  assert.deepEqual(paragraph.runs.map(r => [r.style.fontSize, r.style.bold, r.style.underline, r.style.strike]), [[11, false, false, false], [20, false, false, false]]);
  assert.deepEqual(records(s, "run.fontSize").map(p => p.level), ["run", "layout-placeholder"]);
  assert.equal(records(s, "run.bold")[1].level, "slide-list-style");
  assert.equal(records(s, "run.underline")[0].level, "paragraph");
  assert.equal(records(s, "paragraph.bullet")[0].level, "slide-list-style"); s.dispose();
});

test("level defaults apply to the selected level without borrowing a different level or a parent prompt paragraph", async () => {
  const f = textStyleFixture();
  const levels = Array.from({ length: 9 }, (_, i) => `<a:lvl${i + 1}pPr algn="${i % 2 ? "r" : "l"}"><a:defRPr sz="${(i + 10) * 100}"/></a:lvl${i + 1}pPr>`).join("");
  f.set("deck/layout.xml", "sldLayout", inheritedShape({ text: textBody({ list: `<a:defPPr><a:defRPr b="0"/></a:defPPr>${levels}` }) }));
  setLocal(f, textBody({ paragraphs: Array.from({ length: 9 }, (_, i) => `<a:p><a:pPr lvl="${i}"/><a:r><a:t>L${i}</a:t></a:r></a:p>`).join("") }));
  const s = await open(f);
  assert.deepEqual(projected(s).content.map(p => p.runs[0].style.fontSize), [10,11,12,13,14,15,16,17,18]);
  assert.deepEqual(projected(s).content.map(p => p.style.align), ["left","right","left","right","left","right","left","right","left"]);
  assert.ok(projected(s).content.every(p => p.runs[0].style.bold === false)); s.dispose();
});

test("unsupported explicit character declarations block older values while independent fields keep inheriting", async () => {
  for (const [markup, property, forbidden] of [
    ['<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>', "run.color", "445566"],
    ['<a:solidFill><a:srgbClr val="FFFFFF"><a:hueOff val="25000"/></a:srgbClr></a:solidFill>', "run.color", "445566"],
    ['<a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="100001"/></a:srgbClr></a:solidFill>', "run.color", "445566"],
    ['<a:noFill/>', "run.color", "445566"],
    ['<a:latin typeface="+mn-lt"/>', "run.fontFamily", "Parent Font"],
    ['<a:latin typeface=""/>', "run.fontFamily", "Parent Font"],
  ]) {
    const f = textStyleFixture();
    setLocal(f, textBody({ paragraphs: `<a:p><a:r><a:rPr>${markup}</a:rPr><a:t>Blocked</a:t></a:r></a:p>` }));
    const s = await open(f), run = projected(s).content[0].runs[0];
    assert.notDeepEqual(run.style[property.slice(4)], forbidden);
    assert.equal(run.style.fontSize, 20); assert.equal(records(s, property)[0].status, "unsupported");
    assert.equal(records(s, property)[0].value, undefined); assert.equal(records(s, property)[0].level, "run");
    s.dispose();
  }
  for (const attr of ['u="dbl"', 'strike="dblStrike"']) {
    const f = textStyleFixture(); setLocal(f, textBody({ paragraphs: `<a:p><a:r><a:rPr ${attr}/><a:t>Unsupported</a:t></a:r></a:p>` }));
    const s = await open(f), property = attr.startsWith("u") ? "run.underline" : "run.strike";
    assert.equal(records(s, property)[0].status, "unsupported"); assert.equal(projected(s).content[0].runs[0].style[property.slice(4)], false); s.dispose();
  }
});

test("a supported local value can override an unsupported ancestor blocker without retaining the old source value", async () => {
  const f = textStyleFixture();
  f.set("deck/layout.xml", "sldLayout", inheritedShape({ text: textBody({ list: '<a:lvl1pPr><a:defRPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:defRPr></a:lvl1pPr>' }) }));
  setLocal(f, textBody({ paragraphs: '<a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="abcdef"/></a:solidFill></a:rPr><a:t>Local</a:t></a:r><a:r><a:t>Blocked</a:t></a:r></a:p>' }));
  const s = await open(f);
  assert.equal(projected(s).content[0].runs[0].style.color, "ABCDEF");
  assert.notEqual(projected(s).content[0].runs[1].style.color, "112233");
  assert.deepEqual(records(s, "run.color").map(p => p.status), ["mapped", "unsupported"]);
  assert.equal(records(s, "run.color")[1].status, "unsupported");
  assert.match(records(s, "run.color")[1].source.path, /schemeClr/); s.dispose();
});

test("unsupported frame/paragraph choices mask old fields and carry exact source locations", async () => {
  const f = textStyleFixture();
  setLocal(f, textBody({ frame: 'lIns="-1" anchor="just"',
    paragraphs: '<a:p><a:pPr algn="dist" indent="12700"><a:lnSpc><a:spcPts val="2400"/></a:lnSpc><a:spcBef><a:spcPct val="200000"/></a:spcBef><a:buBlip/></a:pPr><a:r><a:t>Unsupported choices</a:t></a:r></a:p>' }));
  const s = await open(f), text = projected(s), paragraph = text.content[0];
  assert.equal(text.frame.margin.left, 7); assert.equal(text.frame.verticalAlign, "top");
  assert.equal(paragraph.style.align, "left"); assert.equal(paragraph.style.hanging, 0);
  assert.equal(paragraph.style.spaceBefore, 0); assert.equal(paragraph.style.lineSpacing, 1);
  assert.deepEqual(paragraph.style.bullet, { type: "none" });
  for (const property of ["frame.margin.left","frame.verticalAlign","paragraph.align","paragraph.hanging","paragraph.lineSpacing","paragraph.spaceBefore","paragraph.bullet"]) {
    const [row] = records(s, property); assert.equal(row.status, "unsupported", property); assert.equal(row.source.part, "/deck/z.xml");
  }
  assert.match(records(s, "frame.margin.left")[0].source.path, /bodyPr\[1\]\/@lIns$/); s.dispose();
});

test("empty insertion paragraphs and breaks use the cascade; end properties never restyle preceding text", async () => {
  const f = textStyleFixture();
  setLocal(f, textBody({ paragraphs: '<a:p><a:r><a:t>Visible</a:t></a:r><a:br><a:rPr b="0"/></a:br><a:endParaRPr sz="9900"/></a:p><a:p><a:endParaRPr sz="1200" b="0"/></a:p><a:p/>' }));
  const s = await open(f), paragraphs = projected(s).content;
  assert.deepEqual(paragraphs[0].runs.map(r => [r.text, r.style.fontSize, r.style.bold]), [["Visible",20,true],["\n",20,false]]);
  assert.equal(paragraphs[1].runs[0].style.fontSize, 12); assert.equal(paragraphs[1].runs[0].style.bold, false);
  assert.equal(paragraphs[2].runs[0].style.fontSize, 20);
  assert.equal(records(s, "run.fontSize")[2].level, "end-paragraph"); s.dispose();
});

test("title/body/other master text styles are selected independently from matching individual master placeholders", async () => {
  for (const [type, expected] of [["title",30], ["ctrTitle",30], ["body",18], ["obj",18], ["ftr",14], ["dt",14], ["sldNum",14]]) {
    const f = textStyleFixture();
    f.setMaster("", '<a:lvl1pPr><a:defRPr sz="3000"/></a:lvl1pPr>');
    f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace("<p:otherStyle/>", '<p:otherStyle><a:lvl1pPr><a:defRPr sz="1400"/></a:lvl1pPr></p:otherStyle>');
    f.set("deck/layout.xml", "sldLayout", "");
    f.set("deck/z.xml", "sld", inheritedShape({ ph: `type="${type}"`, xfrm: placeholderTransform(), preset: "rect", text: textBody() }));
    const s = await open(f);
    assert.equal(projected(s).content[0].runs[0].style.fontSize, expected, type);
    assert.equal(records(s, "run.fontSize")[0].level, "master-text-style");
    assert.match(records(s, "run.fontSize")[0].source.path, /p:txStyles/); s.dispose();
  }
});

test("plain text uses presentation defaultTextStyle over master otherStyle; placeholders keep their own master style", async () => {
  const f = textStyleFixture();
  f.parts["deck/main.data"] = f.parts["deck/main.data"].replace("</p:presentation>", '<p:defaultTextStyle><a:lvl1pPr algn="r"><a:defRPr sz="1600"/></a:lvl1pPr></p:defaultTextStyle></p:presentation>').replace('xmlns:r=', `xmlns:a="${f.A}" xmlns:r=`);
  f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace("<p:otherStyle/>", '<p:otherStyle><a:lvl1pPr><a:defRPr sz="1400" b="1"/></a:lvl1pPr></p:otherStyle>');
  f.set("deck/z.xml", "sld", shape({ text: textBody() }));
  const s = await open(f), ir = normalizePresentation(s.document);
  assert.equal(ir.slides[0].elements[0].text.content[0].runs[0].style.fontSize, 16);
  assert.equal(ir.slides[0].elements[0].text.content[0].runs[0].style.bold, true);
  assert.equal(ir.slides[1].elements[0].text.content[0].runs[0].style.fontSize, 20);
  assert.equal(records(s, "run.fontSize")[0].level, "presentation-default"); s.dispose();
});

test("supported automatic numbering retains style and explicit start value", async () => {
  const f = textStyleFixture();
  setLocal(f, textBody({ paragraphs: '<a:p><a:pPr><a:buAutoNum type="alphaUcPeriod" startAt="7"/></a:pPr><a:r><a:t>Numbered</a:t></a:r></a:p>' }));
  const s = await open(f);
  assert.deepEqual(projected(s).content[0].style.bullet, { type: "number", style: "alphaUpperPeriod", startAt: 7 });
  assert.equal(records(s, "paragraph.bullet")[0].level, "paragraph");
  s.dispose();
});

test("theme color references resolve through source color maps and retain every palette dependency", async () => {
  const f = setTextTheme(textStyleFixture());
  f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace("</p:sldMaster>", '<p:clrMap bg1="lt1" tx1="accent1" bg2="lt2" tx2="dk1" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>');
  setLocal(f, textBody({ paragraphs: '<a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="tx1"><a:alpha val="100000"/></a:schemeClr></a:solidFill></a:rPr><a:t>Theme color</a:t></a:r></a:p>' }));
  const s = await open(f), text = projected(s), [color] = records(s, "run.color");
  assert.equal(text.content[0].runs[0].style.color, "123456");
  assert.equal(color.level, "run"); assert.equal(color.status, "mapped");
  assert.match(color.source.path, /a:schemeClr\[1\]\/@val$/);
  assert.deepEqual(color.resolvedFrom.map(source => source.part), ["/deck/z.xml", "/deck/master.xml", "/deck/theme.xml", "/deck/z.xml"]);
  assert.match(color.resolvedFrom[3].path, /a:alpha/);
  assert.match(color.resolvedFrom[1].path, /@tx1$/); assert.match(color.resolvedFrom[2].path, /a:accent1/);
  assert.equal(s.report.findings.some(finding => finding.source.part === "/deck/z.xml" && /schemeClr|alpha/.test(finding.source.path ?? "") && /not interpreted|not mapped/.test(finding.reason)), false);
  s.dispose();
});

test("slide color-map override wins over master mapping and masterClrMapping delegates to the master", async () => {
  const f = setTextTheme(textStyleFixture());
  f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace("</p:sldMaster>", '<p:clrMap bg1="lt1" tx1="accent1" bg2="lt2" tx2="dk1" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>');
  const mapping = '<a:overrideClrMapping bg1="lt1" tx1="accent2" bg2="lt2" tx2="dk1" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
  setLocal(f, textBody({ paragraphs: '<a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:rPr><a:t>Slide override</a:t></a:r></a:p>' }));
  f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace("</p:sld>", `<p:clrMapOvr>${mapping}</p:clrMapOvr></p:sld>`);
  const s = await open(f);
  assert.equal(projected(s).content[0].runs[0].style.color, "654321");
  assert.match(records(s, "run.color")[0].resolvedFrom[1].path, /@tx1$/); s.dispose();

  const delegated = setTextTheme(textStyleFixture());
  delegated.parts["deck/master.xml"] = delegated.parts["deck/master.xml"].replace("</p:sldMaster>", '<p:clrMap bg1="lt1" tx1="accent1" bg2="lt2" tx2="dk1" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>');
  setLocal(delegated, textBody({ paragraphs: '<a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:rPr><a:t>Master map</a:t></a:r></a:p>' }));
  delegated.parts["deck/z.xml"] = delegated.parts["deck/z.xml"].replace("</p:sld>", '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>');
  const d = await open(delegated);
  assert.equal(projected(d).content[0].runs[0].style.color, "123456"); d.dispose();

  const inheritedLayoutMap = setTextTheme(textStyleFixture());
  inheritedLayoutMap.parts["deck/master.xml"] = inheritedLayoutMap.parts["deck/master.xml"].replace("</p:sldMaster>", '<p:clrMap bg1="lt1" tx1="accent1" bg2="lt2" tx2="dk1" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>');
  inheritedLayoutMap.parts["deck/layout.xml"] = inheritedLayoutMap.parts["deck/layout.xml"].replace("</p:sldLayout>",
    `<p:clrMapOvr>${mapping}</p:clrMapOvr></p:sldLayout>`);
  setLocal(inheritedLayoutMap, textBody({ paragraphs: '<a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:rPr><a:t>Layout color map</a:t></a:r></a:p>' }));
  const l = await open(inheritedLayoutMap);
  assert.equal(projected(l).content[0].runs[0].style.color, "654321");
  assert.equal(records(l, "run.color")[0].resolvedFrom[1].part, "/deck/layout.xml");
  l.dispose();

  const slideOverridesInherited = setTextTheme(textStyleFixture());
  slideOverridesInherited.parts["deck/master.xml"] = slideOverridesInherited.parts["deck/master.xml"].replace("</p:sldMaster>", '<p:clrMap bg1="lt1" tx1="accent1" bg2="lt2" tx2="dk1" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>');
  slideOverridesInherited.parts["deck/layout.xml"] = slideOverridesInherited.parts["deck/layout.xml"].replace(layoutList,
    '<a:lvl1pPr><a:defRPr><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:lvl1pPr>');
  setLocal(slideOverridesInherited, textBody());
  slideOverridesInherited.parts["deck/z.xml"] = slideOverridesInherited.parts["deck/z.xml"].replace("</p:sld>", `<p:clrMapOvr>${mapping}</p:clrMapOvr></p:sld>`);
  const si = await open(slideOverridesInherited);
  assert.equal(projected(si).content[0].runs[0].style.color, "654321");
  assert.equal(records(si, "run.color")[0].source.part, "/deck/layout.xml");
  assert.ok(records(si, "run.color")[0].resolvedFrom.some(source => source.part === "/deck/z.xml"));
  si.dispose();
});

test("major/minor Latin theme fonts resolve to concrete family names with theme provenance", async () => {
  const f = setTextTheme(textStyleFixture());
  setLocal(f, textBody({ paragraphs: '<a:p><a:r><a:rPr><a:latin typeface="+mj-lt"/></a:rPr><a:t>Major</a:t></a:r><a:r><a:rPr><a:latin typeface="+mn-lt"/></a:rPr><a:t>Minor</a:t></a:r></a:p>' }));
  const s = await open(f), runs = projected(s).content[0].runs;
  assert.deepEqual(runs.map(run => run.style.fontFamily), ["Major Theme Font", "Minor Theme Font"]);
  const fonts = records(s, "run.fontFamily");
  assert.deepEqual(fonts.map(info => info.resolvedFrom[0].part), ["/deck/theme.xml", "/deck/theme.xml"]);
  assert.match(fonts[0].resolvedFrom[0].path, /a:majorFont/);
  s.dispose();
});

test("part theme overrides replace the owning theme's selected color and Latin font slots", async () => {
  const f = setTextTheme(textStyleFixture(), { accent1: "AA3300", major: "Base Major" });
  const base = f.parts["deck/theme.xml"];
  const override = base.replace("<a:theme ", "<a:themeOverride ")
    .replace("</a:themeElements></a:theme>", "</a:themeOverride>").replace("<a:themeElements>", "")
    .replace("AA3300", "BB4400").replace("Base Major", "Override Major");
  f.parts["deck/themeOverride.xml"] = override;
  f.types["deck/themeOverride.xml"] = "themeOverride";
  f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), f.rel("override", "themeOverride", "themeOverride.xml")]);
  setLocal(f, textBody({ paragraphs: '<a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:rPr><a:t>Overridden theme</a:t></a:r></a:p>' }));
  const s = await open(f), run = projected(s).content[0].runs[0];
  assert.equal(run.style.color, "BB4400"); assert.equal(run.style.fontFamily, "Override Major");
  assert.ok(records(s, "run.color")[0].resolvedFrom.some(source => source.part === "/deck/themeOverride.xml"));
  assert.ok(records(s, "run.fontFamily")[0].resolvedFrom.some(source => source.part === "/deck/themeOverride.xml"));
  s.dispose();
});

test("two masters resolve the same theme slots and font roles against their own theme parts", async () => {
  const f = setTextTheme(textStyleFixture(), { accent1: "AA1100", major: "Theme One Font" });
  const themedList = '<a:lvl1pPr><a:defRPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr>';
  f.parts["deck/layout.xml"] = f.parts["deck/layout.xml"].replace(layoutList, themedList);
  f.parts["deck/main.data"] = f.parts["deck/main.data"].replace("</p:sldMasterIdLst>", '<p:sldMasterId id="2147483650" r:id="m2"/></p:sldMasterIdLst>');
  f.parts["deck/_rels/main.data.rels"] = relationships([f.rel("a", "slide", "a.xml"), f.rel("m", "slideMaster", "master.xml"), f.rel("z", "slide", "z.xml"), f.rel("m2", "slideMaster", "master2.xml")]);
  f.parts["deck/layout2.xml"] = f.parts["deck/layout.xml"];
  f.parts["deck/master2.xml"] = f.parts["deck/master.xml"];
  f.parts["deck/theme2.xml"] = f.parts["deck/theme.xml"].replace("AA1100", "00BB22").replace("Theme One Font", "Theme Two Font");
  f.parts["deck/_rels/layout2.xml.rels"] = relationships([f.rel("master", "slideMaster", "master2.xml")]);
  f.parts["deck/_rels/master2.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout2.xml"), f.rel("theme", "theme", "theme2.xml")]);
  f.parts["deck/_rels/a.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout2.xml")]);
  for (const [target, original] of [["deck/layout2.xml", "deck/layout.xml"], ["deck/master2.xml", "deck/master.xml"], ["deck/theme2.xml", "deck/theme.xml"]]) f.types[target] = f.types[original];
  const s = await open(f), runs = normalizePresentation(s.document).slides.map(slide => slide.elements[0].text.content[0].runs[0].style);
  assert.deepEqual(runs.map(run => [run.color, run.fontFamily]), [["AA1100", "Theme One Font"], ["00BB22", "Theme Two Font"]]);
  const colors = s.report.textStyleInheritance.properties.filter(info => info.property === "run.color");
  assert.deepEqual(colors.map(color => color.resolvedFrom.at(-1).part), ["/deck/theme.xml", "/deck/theme2.xml"]);
  s.dispose();
});

test("theme references with unsupported modifiers block inherited colors and script-specific theme fonts", async () => {
  const f = setTextTheme(textStyleFixture());
  setLocal(f, textBody({ paragraphs: '<a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="accent1"><a:hueMod val="50000"/></a:schemeClr></a:solidFill><a:ea typeface="+mn-ea"/></a:rPr><a:t>Unsupported theme variants</a:t></a:r></a:p>' }));
  const s = await open(f), run = projected(s).content[0].runs[0];
  assert.notEqual(run.style.color, "445566");
  assert.equal(records(s, "run.color")[0].status, "unsupported");
  assert.equal(records(s, "run.color")[0].level, "run");
  assert.equal(records(s, "run.fontFamily")[0].status, "unsupported");
  assert.notEqual(run.style.fontFamily, "Parent Font");
  assert.match(records(s, "run.color")[0].reason, /Unsupported color transform hueMod/);
  s.dispose();
});

test("ambiguous parent placeholders never contribute list/body properties, while the document master text style remains available", async () => {
  const f = textStyleFixture();
  f.set("deck/layout.xml", "sldLayout", inheritedShape({ text: textBody({ list: layoutList }) }) + inheritedShape({ id: "3", text: textBody({ list: masterList }) }));
  f.set("deck/z.xml", "sld", inheritedShape({ xfrm: placeholderTransform(), preset: "rect", text: textBody() }));
  const s = await open(f);
  assert.equal(projected(s).content[0].runs[0].style.fontSize, 30);
  assert.equal(projected(s).frame.margin.left, 7);
  assert.ok(s.report.textStyleInheritance.properties.filter(p => p.target.part === "/deck/z.xml").every(p => !["master-placeholder","layout-placeholder"].includes(p.level))); s.dispose();
});

test("multiple masters, shared templates and later local edits do not cross-contaminate style values or provenance", async () => {
  const f = textStyleFixture();
  f.parts["deck/main.data"] = f.parts["deck/main.data"].replace("</p:sldMasterIdLst>", '<p:sldMasterId id="2147483650" r:id="m2"/></p:sldMasterIdLst>');
  f.parts["deck/_rels/main.data.rels"] = relationships([f.rel("a", "slide", "a.xml"), f.rel("m", "slideMaster", "master.xml"), f.rel("z", "slide", "z.xml"), f.rel("m2", "slideMaster", "master2.xml")]);
  f.parts["deck/layout2.xml"] = f.parts["deck/layout.xml"].replace('sz="2000"', 'sz="1700"');
  f.parts["deck/master2.xml"] = f.parts["deck/master.xml"].replace("Parent Font", "Other Master Font");
  f.parts["deck/theme2.xml"] = f.parts["deck/theme.xml"];
  f.parts["deck/_rels/layout2.xml.rels"] = relationships([f.rel("master", "slideMaster", "master2.xml")]);
  f.parts["deck/_rels/master2.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout2.xml"), f.rel("theme", "theme", "theme2.xml")]);
  f.parts["deck/_rels/a.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout2.xml")]);
  for (const [target, source] of [["deck/layout2.xml","deck/layout.xml"],["deck/master2.xml","deck/master.xml"],["deck/theme2.xml","deck/theme.xml"]]) f.types[target] = f.types[source];
  const s = await open(f), ir = normalizePresentation(s.document);
  assert.deepEqual(ir.slides.map(slide => [slide.elements[0].text.content[0].runs[0].style.fontSize, slide.elements[0].text.content[0].runs[0].style.fontFamily]), [[20,"Parent Font"],[17,"Other Master Font"]]);
  assert.equal(records(s, "run.fontSize", "/deck/a.xml")[0].source.part, "/deck/layout2.xml");
  const old = s.document.slides[0].removeElement(s.document.slides[0].elements[0].id), next = structuredClone(old);
  next.text.content[0].runs[0].style.fontSize = 42; s.document.slides[0].addElement(next);
  assert.equal(normalizePresentation(s.document).slides[1].elements[0].text.content[0].runs[0].style.fontSize, 17);
  assert.equal(records(s, "run.fontSize")[0].value, 20, "report remains an immutable source projection snapshot"); s.dispose();
});

test("inherited unsupported properties are located once per shared source style rather than once per page", async () => {
  const f = textStyleFixture();
  f.parts["deck/layout.xml"] = f.parts["deck/layout.xml"].replace('<a:bodyPr ', '<a:bodyPr rtlCol="1" ').replace("<a:defRPr ", '<a:defRPr baseline="100" ');
  const s = await open(f);
  const warnings = s.report.findings.filter(f => f.feature === "drawing.inherited-text.unmapped-property");
  assert.equal(warnings.filter(f => f.source.path.endsWith("/@rtlCol")).length, 1);
  assert.equal(warnings.filter(f => f.source.path.endsWith("/@baseline")).length, 1);
  assert.ok(warnings.every(f => f.source.part === "/deck/layout.xml")); s.dispose();
});

test("inherited style singleton/range errors fail before publishing; unknown extension islands remain opaque", async () => {
  for (const markup of [
    '<a:lvl1pPr/><a:lvl1pPr/>',
    '<a:lvl1pPr><a:defRPr b="maybe"/></a:lvl1pPr>',
    '<a:lvl1pPr><a:defRPr sz="-1"/></a:lvl1pPr>',
    '<a:lvl1pPr><a:buNone/><a:buChar char="*"/></a:lvl1pPr>',
    '<a:lvl1pPr><a:lnSpc><a:spcPts val="100"/><a:spcPct val="100000"/></a:lnSpc></a:lvl1pPr>',
  ]) {
    const f = textStyleFixture(); f.set("deck/layout.xml", "sldLayout", inheritedShape({ text: textBody({ list: markup }) }));
    await assert.rejects(open(f), code("import-invalid-package"));
  }
  const f = textStyleFixture();
  f.parts["deck/layout.xml"] = f.parts["deck/layout.xml"].replace("</a:lstStyle>", '<a:extLst><a:ext uri="urn:opaque"><a:lvl1pPr><a:defRPr sz="-1"/></a:lvl1pPr></a:ext></a:extLst></a:lstStyle>');
  const s = await open(f); assert.equal(projected(s).content[0].runs[0].style.fontSize, 20);
  assert.ok(s.report.findings.some(f => f.source.path?.includes("extLst"))); s.dispose();
});

test("invalid paragraph levels, scalar text and local style choices cannot bypass parsing validation", async () => {
  for (const paragraph of [
    '<a:p><a:pPr lvl="9"/><a:r><a:t>T</a:t></a:r></a:p>',
    '<a:p><a:pPr lvl="-1"/><a:r><a:t>T</a:t></a:r></a:p>',
    '<a:p><a:r><a:t>before<a:r/>after</a:t></a:r></a:p>',
  ]) {
    const f = textStyleFixture(); setLocal(f, textBody({ paragraphs: paragraph }));
    await assert.rejects(open(f), code("import-invalid-package"));
  }
});

test("style provenance budget has exact boundaries, captured options, cancellation and independent session state", async () => {
  assert.equal(DEFAULT_SESSION_LIMITS.maxTextStyleProperties, 500000);
  const f = textStyleFixture(), source = encodeStructure(f), baseline = await open(f);
  const count = baseline.report.textStyleInheritance.properties.length;
  assert.ok(count > 20); baseline.dispose();
  const options = { projection: "content", sessionLimits: { maxTextStyleProperties: count } };
  const pending = importPptx(source, options); options.sessionLimits.maxTextStyleProperties = 1;
  const a = await pending, b = await open(f);
  assert.deepEqual(a.report.textStyleInheritance, b.report.textStyleInheritance); a.dispose(); b.dispose();
  await assert.rejects(open(f, { sessionLimits: { maxTextStyleProperties: count - 1 } }), code("import-limit-exceeded"));
  await assert.rejects(open(f, { sessionLimits: { maxReportFindings: 1 } }), code("import-limit-exceeded"));
  for (const value of [0,-1,1.5,NaN,Infinity]) await assert.rejects(open(f, { sessionLimits: { maxTextStyleProperties: value } }), RangeError);
  const controller = new AbortController(), aborted = open(f, { signal: controller.signal }); controller.abort();
  await assert.rejects(aborted, code("operation-aborted"));
  const structure = await open(f, { projection: "structure", sessionLimits: { maxTextStyleProperties: 1 } });
  assert.equal(structure.report.textStyleInheritance, undefined); structure.dispose();
});

test("notes reuse local list styles and explicit blockers without importing notes-master prompt/default styles", async () => {
  const text = textBody({ list: '<a:lvl1pPr><a:defRPr sz="1500" b="1" u="sng"/></a:lvl1pPr>',
    paragraphs: '<a:p><a:r><a:rPr b="0" u="dbl"/><a:t>Notes local list</a:t></a:r></a:p>' });
  const f = notesFixture(notesShape({ text })), s = await open(f);
  const run = normalizePresentation(s.document).slides[0].notes[0].runs[0];
  assert.equal(run.style.fontSize, 15); assert.equal(run.style.bold, false); assert.equal(run.style.underline, false);
  assert.equal(run.text, "Notes local list");
  assert.ok(s.report.findings.some(f => f.feature === "notes.unmapped-property" && /Unsupported u/.test(f.reason)));
  assert.ok(s.report.textStyleInheritance.properties.every(p => !p.target.part.includes("notes/"))); s.dispose();
});

test("table cell local list styles share the parser while table frame ownership and inherited source scope stay separate", async () => {
  const body = textBody({ frame: 'lIns="12700" anchor="b" wrap="none"',
    list: '<a:lvl1pPr algn="r"><a:defRPr sz="1700" b="1"/></a:lvl1pPr>',
    paragraphs: '<a:p><a:r><a:rPr b="0"/><a:t>Table local list</a:t></a:r></a:p>' }).replaceAll("p:txBody", "a:txBody");
  const f = tableFixture(table({ columns: [100], rows: [tableRow([cell({ body })], 50)] })), s = await open(f);
  const tableCell = normalizePresentation(s.document).slides[0].elements[0].rows[0].cells[0];
  assert.equal(tableCell.content[0].style.align, "right"); assert.equal(tableCell.content[0].runs[0].style.fontSize, 17);
  assert.equal(tableCell.content[0].runs[0].style.bold, false);
  assert.equal(tableCell.style.margin.left, 1); assert.equal(tableCell.style.verticalAlign, "bottom");
  assert.ok(s.report.findings.some(f => f.feature === "drawing.unmapped-property" && f.source.path.endsWith("/@wrap")));
  assert.deepEqual(s.report.textStyleInheritance.properties, []); s.dispose();
});
