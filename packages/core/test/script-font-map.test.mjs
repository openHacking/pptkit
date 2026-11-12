import test from "node:test";
import assert from "node:assert/strict";
import { createPresentation, normalizePresentation, validatePresentation, resolveTextFontSegments, TEXT_FONT_SCRIPT_CODES } from "../dist/index.js";

const box = { x: 0, y: 0, width: 300, height: 100 };
const content = style => [{ runs: [{ text: "A中عא", style }] }];
const theme = normalizePresentation(createPresentation({ theme: { fonts: { heading: "Heading", body: "Body" } } })).theme;
const map = { Latn: "Latin", Hans: "Simplified", Hant: "Traditional", Hani: "Han", Jpan: "Japanese", Kore: "Korean", Arab: "Arabic", Hebr: "Hebrew" };
const choose = (text, language, scriptFontFamilies = map) => resolveTextFontSegments(text, { fontFamily: "Default", eastAsianFontFamily: "East", complexScriptFontFamily: "Complex", scriptFontFamilies, language }, theme);

test("script maps detach across snapshots, duplicate, shape, notes, JSON and normalized output", () => {
  const input = { scriptFontFamilies: { Arab: { theme: "heading" }, Hebr: "Hebrew" } }, p = createPresentation(), s = p.addSlide({ notes: content(input) });
  const e = s.addElement({ type: "text", box, content: content(input) });
  s.addElement({ type: "shape", shape: "rect", box, text: { content: content(input) } });
  input.scriptFontFamilies.Arab.theme = "body";
  assert.deepEqual(e.content[0].runs[0].style.scriptFontFamilies.Arab, { theme: "heading" });
  assert.throws(() => { e.content[0].runs[0].style.scriptFontFamilies.Hebr = "Changed"; }, TypeError);
  s.duplicateElement(e.id); const n = normalizePresentation(p);
  assert.deepEqual(validatePresentation(p), []); assert.deepEqual(JSON.parse(JSON.stringify(n)), n);
  for (const run of [n.slides[0].notes[0].runs[0], ...n.slides[0].elements.map(e => (e.type === "shape" ? e.text : e).content[0].runs[0])]) assert.deepEqual(run.style.scriptFontFamilies, { Arab: { theme: "heading" }, Hebr: "Hebrew" });
  n.slides[0].elements[0].content[0].runs[0].style.scriptFontFamilies.Arab.theme = "body";
  assert.deepEqual(e.content[0].runs[0].style.scriptFontFamilies.Arab, { theme: "heading" });
  assert.ok(Object.isFrozen(TEXT_FONT_SCRIPT_CODES)); assert.equal(new Set(TEXT_FONT_SCRIPT_CODES).size, TEXT_FONT_SCRIPT_CODES.length);
});

test("preset and placeholder maps inherit per key, with key and whole-map reset", () => {
  const p = createPresentation({ textStylePresets: { fonts: { run: { scriptFontFamilies: map } } } });
  p.defineSlideLayout({ id: "layout", name: "Layout", placeholders: [{ key: "title", kind: "title", box, textStyle: { run: { scriptFontFamilies: map } } }] });
  const s = p.addSlide({ layoutId: "layout" });
  for (const patch of [{ Arab: "Local", Hebr: null, Hans: undefined }, null, {}]) s.addElement({ type: "text", box, textStylePreset: "fonts", content: content({ scriptFontFamilies: patch }) });
  s.addElement({ type: "text", placeholderKey: "title", content: content({ scriptFontFamilies: { Arab: null } }) });
  const runs = normalizePresentation(p).slides[0].elements.map(e => e.content[0].runs[0]);
  const { Hebr: _removed, ...retained } = map;
  assert.deepEqual(runs[0].style.scriptFontFamilies, { ...retained, Arab: "Local" });
  assert.equal(Object.hasOwn(runs[1].style, "scriptFontFamilies"), false);
  assert.deepEqual(runs[2].style.scriptFontFamilies, map);
  const { Arab: _arab, ...placeholderMap } = map;
  assert.deepEqual(runs[3].style.scriptFontFamilies, placeholderMap);
});

test("regional table merges preserve full reset barriers through later object and local overrides", () => {
  const p = createPresentation({ textStylePresets: { cell: { run: { scriptFontFamilies: { Arab: "Preset Arabic", Hebr: "Preset Hebrew" } } } } });
  p.addSlide().addElement({ type: "table", box, columns: [150, 150], style: {
    wholeTable: { text: { run: { scriptFontFamilies: { Arab: "Region Arabic", Hebr: "Region Hebrew" } } } },
    firstRow: { text: { run: { scriptFontFamilies: null } } },
    topLeft: { text: { run: { scriptFontFamilies: { Arab: "Corner" } } } },
  }, rows: [{ cells: [
    { textStylePreset: "cell", content: content({ scriptFontFamilies: { Arab: null, Latn: "Local Latin" } }) },
    { content: content({ scriptFontFamilies: { Hebr: "Local Hebrew" } }) },
  ] }] });
  const cells = normalizePresentation(p).slides[0].elements[0].rows[0].cells;
  // Cell preset intentionally comes after regions; local null clears only Arab.
  assert.deepEqual(cells[0].content[0].runs[0].style.scriptFontFamilies, { Latn: "Local Latin", Hebr: "Preset Hebrew" });
  assert.deepEqual(cells[1].content[0].runs[0].style.scriptFontFamilies, { Hebr: "Local Hebrew" });
  assert.deepEqual(validatePresentation(p), []);
  p.slides[0].addElement({ type: "table", box, columns: [300], style: {
    wholeTable: { text: { run: { scriptFontFamilies: map } } },
    firstRow: { text: { run: { scriptFontFamilies: null } } },
    topLeft: { text: { run: { scriptFontFamilies: { Arab: "Only Local" } } } },
  }, rows: [{ cells: [{ content: "Reset" }] }] });
  const reset = normalizePresentation(p).slides[0].elements[1].rows[0].cells[0].content[0].runs[0].style;
  assert.deepEqual(reset.scriptFontFamilies, { Arab: "Only Local" }, "whole-table defaults must not resurrect after a whole-map reset followed by a corner map");
});

test("invalid map roots, keys and family values have located diagnostics across owners", () => {
  for (const value of [[], "Arabic", 1, false, { XXXX: "Bad" }, { Arab: "" }, { Arab: "Bad\nFont" }, { Arab: { theme: "unknown" } }, { Arab: { theme: "body", extra: true } }]) {
    const style = { scriptFontFamilies: value }, p = createPresentation({ textStylePresets: { bad: { run: style } } }), s = p.addSlide({ notes: content(style) });
    p.defineSlideLayout({ id: "bad", name: "Bad", placeholders: [{ key: "title", kind: "title", box, textStyle: { run: style } }] });
    s.addElement({ type: "text", box, content: content(style) });
    s.addElement({ type: "shape", shape: "rect", box, text: { content: content(style) } });
    s.addElement({ type: "table", box, columns: [300], rows: [{ cells: [{ content: content(style) }] }], style: { wholeTable: { text: { run: style } } } });
    const issues = validatePresentation(p).filter(d => d.path.includes("scriptFontFamilies"));
    assert.equal(issues.length, 7, JSON.stringify(value)); assert.ok(issues.every(d => d.code === "invalid-script-font-map" || d.code === "invalid-table-style"));
  }
});

test("explicit language components select Han variants without locale maximization", () => {
  const cases = { "zh-CN": "Simplified", "zh-SG": "Simplified", "zh-TW": "Traditional", "zh-HK": "Traditional", "zh-MO": "Traditional", "yue-HK": "Traditional", "zh-Hans-HK": "Simplified", "zh-Hant-CN": "Traditional", ja: "Japanese", ko: "Korean", "zh-Jpan-CN": "Japanese", "zh-Kore-TW": "Korean", zh: "Han", yue: "Han", "zh-Latn-CN": "Han", "en-US": "Han", "bad_locale": "Han", "zh-u-rg-twzzzz": "Han", "zh-x-hant": "Han" };
  for (const [language, family] of Object.entries(cases)) assert.equal(choose("中𠀀", language)[0].family, family, language);
  assert.equal(choose("中", undefined)[0].family, "Han");
  assert.equal(choose("中", "zh-TW", { Hans: "Wrong" })[0].family, "East");
  assert.equal(choose("中", "zh-TW", { Hani: { theme: "heading" } })[0].family, "Heading");
});

test("exact script entries precede language aliases and ordinary slots", () => {
  assert.deepEqual(choose("あカ한ㄅ", "zh-TW", { ...map, Hira: "Hiragana", Hang: "Hangul" }).map(s => [s.text, s.family, s.scriptFont]), [["あ", "Hiragana", "Hira"], ["カ", "Japanese", "Jpan"], ["한", "Hangul", "Hang"], ["ㄅ", "Traditional", "Hant"]]);
  assert.equal(choose("ㄅ", "zh-CN")[0].family, "East");
  assert.deepEqual(choose("AΩЯعא", "en", { ...map, Grek: "Greek", Cyrl: "Cyrillic" }).map(s => s.family), ["Latin", "Greek", "Cyrillic", "Arabic", "Hebrew"]);
});

test("all named script keys select their actual code point controls", () => {
  const controls = { Latn: "A", Grek: "Ω", Cyrl: "Я", Hani: "中", Hira: "あ", Kana: "カ", Hang: "한", Bopo: "ㄅ", Yiii: "ꆈ", Arab: "ع", Hebr: "א", Syrc: "ܐ", Thaa: "ހ", Nkoo: "ߊ", Samr: "\u0800", Mand: "\u0840", Deva: "अ", Beng: "ব", Guru: "ਗ", Gujr: "ગ", Orya: "ଅ", Taml: "அ", Telu: "అ", Knda: "ಅ", Mlym: "അ", Sinh: "අ", Thai: "ก", Laoo: "ກ", Tibt: "ཀ", Mymr: "က", Khmr: "ក", Mong: "ᠠ", Zyyy: "😀" };
  for (const [code, text] of Object.entries(controls)) assert.equal(choose(text, undefined, { [code]: code })[0].scriptFont, code, code);
  assert.deepEqual(new Set([...Object.keys(controls), "Hans", "Hant", "Jpan", "Kore"]), new Set(TEXT_FONT_SCRIPT_CODES));
});

test("map mode preserves marks, UTF16 text, newline context and exact segment budgets", () => {
  const text = "𠀀A\u0301عَ\u200Dب! א\n!😀", segments = choose(text, "zh-CN");
  assert.equal(segments.map(s => s.text).join(""), text); assert.equal(segments[0].end, 2);
  assert.ok(segments.some(s => s.text === "عَ\u200Dب! " && s.scriptFont === "Arab"));
  for (const [i, s] of segments.entries()) { assert.equal(text.slice(s.start, s.end), s.text); assert.equal(s.start, i ? segments[i - 1].end : 0); }
  assert.equal(choose("中、", "zh-TW").at(-1).scriptFont, "Hant");
  assert.equal(choose("、中", "zh-TW")[0].scriptFont, undefined);
  for (const newline of ["\n", "\r", "\u2028", "\u2029"]) assert.deepEqual(choose(`ع${newline}!`, "en").map(s => [s.text, s.family]), [["ع", "Arabic"], [`${newline}!`, "Default"]]);
  assert.equal(choose("", "zh-TW")[0].family, "Default");
  const style = { fontFamily: "Default", scriptFontFamilies: map };
  assert.equal(resolveTextFontSegments("عאع", style, theme, { maxSegments: 3 }).length, 3);
  assert.throws(() => resolveTextFontSegments("عאع", style, theme, { maxSegments: 2 }), /maxSegments/);
  const legacy = { fontFamily: "Legacy" };
  assert.deepEqual(resolveTextFontSegments(text, legacy, theme), resolveTextFontSegments(text, { ...legacy, scriptFontFamilies: {} }, theme));
});
