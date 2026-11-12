import test from "node:test";
import assert from "node:assert/strict";
import { createPresentation, normalizePresentation, validatePresentation, resolveTextFontSegments } from "../dist/index.js";

const box = { x: 0, y: 0, width: 300, height: 100 };
const style = { fontFamily: "Latin", eastAsianFontFamily: "Eastern", complexScriptFontFamily: "Complex" };
const content = runStyle => [{ runs: [{ text: "Latin中文العربية", style: runStyle }] }];
const theme = normalizePresentation(createPresentation({ theme: { fonts: { heading: "Heading", body: "Body" } } })).theme;
const choose = text => resolveTextFontSegments(text, style, theme);

test("script families detach through author snapshots, normalization, duplicate, notes and JSON IR", () => {
  const input = { ...style, eastAsianFontFamily: { theme: "heading" } }, p = createPresentation(), slide = p.addSlide({ notes: content(input) });
  const element = slide.addElement({ type: "text", box, content: content(input) }); input.eastAsianFontFamily.theme = "body";
  assert.deepEqual(element.content[0].runs[0].style.eastAsianFontFamily, { theme: "heading" });
  assert.throws(() => { element.content[0].runs[0].style.eastAsianFontFamily.theme = "body"; }, TypeError);
  slide.duplicateElement(element.id); const normalized = normalizePresentation(p);
  assert.equal(normalized.slides[0].elements.length, 2); assert.deepEqual(validatePresentation(p), []);
  assert.deepEqual(normalized.slides[0].notes[0].runs[0].style.eastAsianFontFamily, { theme: "heading" });
  const json = JSON.parse(JSON.stringify(normalized)); assert.deepEqual(json, normalized);
  normalized.slides[0].elements[0].content[0].runs[0].style.eastAsianFontFamily.theme = "body";
  assert.deepEqual(element.content[0].runs[0].style.eastAsianFontFamily, { theme: "heading" });
});

test("preset and regional families inherit independently, null resets a single slot", () => {
  const p = createPresentation({ textStylePresets: { multi: { run: style } } }), s = p.addSlide();
  p.defineSlideLayout({ id: "script-layout", name: "Script layout", placeholders: [{ key: "title", kind: "title", box, textStyle: { run: style } }] });
  s.addElement({ type: "text", box, textStylePreset: "multi", content: [{ runs: [{ text: "One" }, { text: "Two", style: { fontFamily: "Local Latin", eastAsianFontFamily: null } }] }] });
  s.addElement({ type: "shape", shape: "rect", box, text: { textStylePreset: "multi", content: "Shape" } });
  s.addElement({ type: "table", box, columns: [150, 150], style: { wholeTable: { text: { run: style } }, firstRow: { text: { run: { eastAsianFontFamily: "First" } } } }, rows: [{ cells: [{ content: content({ complexScriptFontFamily: null }) }, { content: "Other" }] }] });
  const [text, shape, table] = normalizePresentation(p).slides[0].elements;
  assert.equal(text.content[0].runs[0].style.eastAsianFontFamily, "Eastern");
  const reset = text.content[0].runs[1].style; assert.equal(reset.fontFamily, "Local Latin"); assert.equal(Object.hasOwn(reset, "eastAsianFontFamily"), false); assert.equal(reset.complexScriptFontFamily, "Complex");
  assert.equal(shape.text.content[0].runs[0].style.eastAsianFontFamily, "Eastern");
  assert.equal(table.rows[0].cells[0].content[0].runs[0].style.eastAsianFontFamily, "First"); assert.equal(table.rows[0].cells[0].content[0].runs[0].style.complexScriptFontFamily, undefined);
  assert.equal(table.rows[0].cells[1].content[0].runs[0].style.complexScriptFontFamily, "Complex");
  const bound = p.addSlide({ layoutId: "script-layout" });
  bound.addElement({ type: "text", placeholderKey: "title", content: content({ complexScriptFontFamily: null }) });
  const inherited = normalizePresentation(p).slides[1].elements[0].content[0].runs[0].style;
  assert.equal(inherited.eastAsianFontFamily, "Eastern"); assert.equal(inherited.complexScriptFontFamily, undefined);
});

test("invalid script families produce located diagnostics across direct/preset/shape/cell/notes owners", () => {
  for (const field of ["eastAsianFontFamily", "complexScriptFontFamily"]) for (const value of ["", "   ", "Font\nName", "Font\u007FName", 3, false, [], {}, { theme: "unknown" }, { theme: "body", extra: true }]) {
    const input = { [field]: value }, p = createPresentation({ textStylePresets: { bad: { run: input } } }), s = p.addSlide({ notes: content(input) });
    s.addElement({ type: "text", box, content: content(input) });
    s.addElement({ type: "shape", shape: "rect", box, text: { content: content(input) } });
    s.addElement({ type: "table", box, columns: [300], rows: [{ cells: [{ content: content(input) }] }] });
    const issues = validatePresentation(p).filter(d => d.path.endsWith(`.${field}`)); assert.equal(issues.length, 5, `${field}/${JSON.stringify(value)}`); assert.ok(issues.every(d => d.code === "invalid-font-family"));
  }
  const p = createPresentation(); p.addSlide().addElement({ type: "table", box, columns: [300], rows: [{ cells: [{ content: "Cell" }] }], style: { wholeTable: { text: { run: { eastAsianFontFamily: [] } } } } });
  assert.ok(validatePresentation(p).some(d => d.path.endsWith(".style.wholeTable.text.run.eastAsianFontFamily")));
});

test("omitted slots preserve old normalized IR and single-family text behavior", () => {
  const p = createPresentation(); p.addSlide().addElement({ type: "text", box, content: content({ fontFamily: "Only" }) });
  const run = normalizePresentation(p).slides[0].elements[0].content[0].runs[0];
  assert.equal(Object.hasOwn(run.style, "eastAsianFontFamily"), false); assert.equal(Object.hasOwn(run.style, "complexScriptFontFamily"), false);
  assert.deepEqual(resolveTextFontSegments(run.text, run.style, theme), [{ text: run.text, start: 0, end: run.text.length, family: "Only", script: "default" }]);
});

test("named Unicode script controls select actual families including supplementary Han", () => {
  const eastern = ["中", "あ", "カ", "한", "ㄅ", "ꆈ", "𠀀"];
  const complex = ["ع", "א", "ܐ", "ހ", "ߊ", "\u0800", "\u0840", "अ", "ব", "ਗ", "ગ", "ଅ", "அ", "అ", "ಅ", "അ", "අ", "ก", "ກ", "ཀ", "က", "ក", "ᠠ"];
  for (const c of eastern) assert.equal(choose(c)[0].family, "Eastern", c);
  for (const c of complex) assert.equal(choose(c)[0].family, "Complex", c);
  for (const c of ["A", "é", "Ω", "Я", "1", "😀"]) assert.equal(choose(c)[0].family, "Latin", c);
});

test("marks/joiners remain with preceding font, punctuation and separators have declared context", () => {
  assert.deepEqual(choose("A\u0301中\uFE00ع\u064E\u200Dب").map(s => [s.text, s.family]), [["A\u0301", "Latin"], ["中\uFE00", "Eastern"], ["ع\u064E\u200Dب", "Complex"]]);
  assert.deepEqual(choose("中! A\nع? B").map(s => [s.text, s.family]), [["中! ", "Eastern"], ["A\n", "Latin"], ["ع? ", "Complex"], ["B", "Latin"]]);
  assert.deepEqual(choose("😀👩‍💻").map(s => s.family), ["Latin"]);
});

test("offsets reconstruct original UTF-16 without splitting code points or mutating inputs", () => {
  const input = "𠀀A\u0301عَ😀中", frozen = Object.freeze({ ...style }), segments = resolveTextFontSegments(input, frozen, theme);
  assert.equal(segments.map(s => s.text).join(""), input);
  assert.equal(segments[0].end, 2);
  for (const [i, s] of segments.entries()) { assert.equal(input.slice(s.start, s.end), s.text); assert.equal(s.start, i === 0 ? 0 : segments[i - 1].end); }
  assert.deepEqual(frozen, style);
  const themed = resolveTextFontSegments("A中ع", { fontFamily: { theme: "body" }, eastAsianFontFamily: { theme: "heading" } }, theme);
  assert.deepEqual(themed.map(s => s.family), ["Body", "Heading", "Body"]);
});

test("empty insertion and exact segment budgets reject excessive alternating text without partial output", () => {
  assert.deepEqual(choose(""), [{ text: "", start: 0, end: 0, family: "Latin", script: "default" }]);
  assert.equal(resolveTextFontSegments("A中ع", style, theme, { maxSegments: 3 }).length, 3);
  assert.throws(() => resolveTextFontSegments("A中ع", style, theme, { maxSegments: 2 }), /maxSegments/);
  for (const maximum of [0, -1, 1.5, NaN, Infinity, Number.MAX_SAFE_INTEGER + 1]) assert.throws(() => resolveTextFontSegments("", style, theme, { maxSegments: maximum }), /positive safe integer/);
  assert.equal(resolveTextFontSegments("中".repeat(100_000), style, theme, { maxSegments: 1 })[0].end, 100_000);
  assert.equal(choose("Retry")[0].text, "Retry");
});
