import test from "node:test";
import assert from "node:assert/strict";
import { transformTextColor } from "../dist/import/color-transforms.js";
import { importPptx, DEFAULT_SESSION_LIMITS } from "../dist/index.js";
import { textBody, textStyleFixture, setTextTheme, addTextThemeOverride } from "./text-style-fixture.mjs";
import { inheritedShape } from "./placeholder-fixture.mjs";
import { encodeStructure } from "./structure-fixture.mjs";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const node = (local, value) => ({
  kind: "element", name: `a:${local}`, local, namespaceUri: A,
  attributes: value === undefined ? [] : [{ name: "val", local: "val", namespaceUri: "", value: String(value) }],
  children: [],
});
const evaluate = (rgb, ...operations) => transformTextColor(rgb, operations.map(([name, value]) => node(name, value)), A);
const colorXml = (name, value, operations = "") => `<a:${name} ${name === "sysClr" ? 'val="windowText" lastClr' : "val"}="${value}">${operations}</a:${name}>`;
const fill = color => `<a:solidFill>${color}</a:solidFill>`;
const run = (color, text = "Transformed") => `<a:r><a:rPr>${fill(color)}</a:rPr><a:t>${text}</a:t></a:r>`;
function fixture({ color, strict = false } = {}) {
  const f = setTextTheme(textStyleFixture(strict));
  if (color) f.set("deck/z.xml", "sld", inheritedShape({ text: textBody({ paragraphs: `<a:p>${run(color)}</a:p>` }) }));
  return f;
}
const open = (f, options = {}) => importPptx(encodeStructure(f), { projection: "content", ...options });
const style = session => session.document.slides[0].elements[0].text.content[0].runs[0].style;
const records = (session, property) => session.report.textStyleInheritance.properties.filter(p => p.target.part === "/deck/z.xml" && p.property === property);

for (const [name, value, expected] of [
  ["lum", 50000, "4080BF"], ["lumMod", 50000, "1A334D"], ["lumOff", 20000, "6699CC"],
  ["sat", 20000, "52667A"], ["satMod", 50000, "4D6680"], ["satOff", -25000, "4D6680"],
]) {
  test(`${name} uses independent HSL reference values, not RGB multiplication`, () => {
    // Expected constants also checked with Python's standard-library colorsys.
    assert.deepEqual(evaluate("336699", [name, value]), { status: "mapped", value: expected, opacity: 1 });
  });
}

test("tint/shade use linear mixing, boundary values and final-only rounding", () => {
  for (const [rgb, operation, value, expected] of [
    ["000000", "tint", 50000, "BCBCBC"], ["FFFFFF", "shade", 50000, "BCBCBC"],
    ["808080", "shade", 50000, "5C5C5C"], ["123456", "tint", 0, "FFFFFF"],
    ["123456", "shade", 0, "000000"], ["123456", "tint", 100000, "123456"],
    ["123456", "shade", 100000, "123456"],
  ]) assert.equal(evaluate(rgb, [operation, value]).value, expected);
  assert.equal(evaluate("FFFFFF", ["shade", 50000], ["shade", 50000]).value, "898989");
});

test("HSL state and transform order survive achromatic intermediate states", () => {
  assert.equal(evaluate("336699", ["sat", 0], ["sat", 50000]).value, "336699");
  assert.equal(evaluate("336699", ["lum", 0], ["lum", 40000]).value, "336699");
  assert.equal(evaluate("336699", ["lumMod", 50000], ["lumOff", 20000]).value, "336699");
  assert.equal(evaluate("336699", ["lumOff", 20000], ["lumMod", 50000]).value, "264D73");
  assert.notEqual(evaluate("336699", ["shade", 50000], ["lumOff", 20000]).value,
    evaluate("336699", ["lumOff", 20000], ["shade", 50000]).value);
});

test("schema signed/unbounded HSL percentages clamp, and alpha has its own schema ranges", () => {
  assert.equal(evaluate("336699", ["lumMod", -10000]).value, "000000");
  assert.equal(evaluate("336699", ["lum", 200000]).value, "FFFFFF");
  assert.equal(evaluate("336699", ["satMod", -10000]).value, "666666");
  assert.equal(evaluate("336699", ["alpha", 50000], ["alphaMod", 50000], ["alphaOff", 10000]).opacity, .35);
  assert.equal(evaluate("336699", ["alphaOff", -100000]).opacity, 0);
  assert.equal(evaluate("336699", ["alpha", 50000], ["alphaMod", 400000]).opacity, 1);
  assert.equal(evaluate("336699", ["alphaOff", -90000], ["alphaOff", 100000]).opacity, 1);
  assert.equal(evaluate("336699", ["alphaOff", 100000], ["alphaOff", -90000]).opacity, .09999999999999998);
});

test("integer and percent lexical forms agree including whitespace and explicit plus signs", () => {
  for (const name of ["lum", "lumMod", "lumOff", "sat", "satMod", "satOff", "alpha", "alphaMod", "alphaOff", "shade", "tint"]) {
    assert.deepEqual(evaluate("336699", [name, 25000]), evaluate("336699", [name, "  +25.0%  "]));
    assert.deepEqual(evaluate("336699", [name, " +25000 "]), evaluate("336699", [name, "25%"]));
  }
});

test("unknown, malformed, foreign and extended transforms never return a partial color", () => {
  for (const invalid of [
    node("hue", 50000), node("tint"), node("tint", "1e4"), node("tint", "12px"),
    node("tint", "NaN"), node("tint", ".5%"), node("tint", "50.%"), node("tint", "50 %"),
    node("tint", "\u00A050%\u00A0"),
    node("tint", -1), node("tint", 100001), node("alpha", -1), node("alpha", 100001),
    node("alphaOff", 100001), node("alphaMod", -1), node("lum", 2147483648),
    node("lum", "999999999999999999999999999999999999999999"),
    { ...node("lum", 50000), namespaceUri: "urn:other" },
    { ...node("lum", 50000), children: [node("alpha", 1)] },
    { ...node("lum", 50000), children: [{ kind: "text", value: "unexpected" }] },
    { ...node("lum", 50000), attributes: [...node("lum", 50000).attributes, { name: "extra", local: "extra", namespaceUri: "", value: "1" }] },
  ]) {
    const result = transformTextColor("336699", [node("shade", 50000), invalid], A);
    assert.equal(result.status, "unsupported", JSON.stringify(invalid));
    assert.equal(result.value, undefined);
    assert.equal(result.opacity, undefined);
  }
});

test("all 8-bit grays survive identity and alpha-only operations without RGB drift", () => {
  for (let value = 0; value <= 255; value++) {
    const rgb = value.toString(16).padStart(2, "0").repeat(3).toUpperCase();
    assert.equal(evaluate(rgb, ["shade", 100000], ["tint", 100000], ["lumMod", 100000]).value, rgb);
    assert.deepEqual(evaluate(rgb, ["alpha", 0]), { status: "mapped", value: rgb, opacity: 0 });
  }
});

test("direct, system-fallback and scheme colors map alpha and ordered transforms in both dialects", async () => {
  for (const strict of [false, true]) for (const name of ["srgbClr", "sysClr", "schemeClr"]) {
    const s = await open(fixture({ strict, color: colorXml(name, name === "schemeClr" ? "accent1" : "123456", '<a:lumMod val="50%"/><a:alpha val="25000"/>') }));
    try {
      assert.equal(style(s).color, "091A2B");
      assert.equal(style(s).colorOpacity, .25);
      for (const property of ["run.color", "run.colorOpacity"]) {
        const [record] = records(s, property);
        assert.equal(record.status, "mapped");
        assert.ok(record.resolvedFrom.some(location => /lumMod/.test(location.path)));
        assert.ok(record.resolvedFrom.some(location => /alpha/.test(location.path)));
        assert.ok(Object.isFrozen(record.resolvedFrom));
      }
      const findings = s.report.findings.filter(f => f.source.part === "/deck/z.xml" && /lumMod|alpha/.test(f.source.path ?? ""));
      assert.deepEqual(findings, [], "interpreted transforms have no unsupported-subtree findings");
    } finally { s.dispose(); }
  }
});

test("palette and reference transform chains compose before quantization, including partial overrides", async () => {
  const f = fixture({ color: colorXml("schemeClr", "accent1", '<a:shade val="50000"/><a:alphaMod val="50000"/>') });
  const scheme = /<a:clrScheme[\s\S]*?<\/a:clrScheme>/.exec(f.parts["deck/theme.xml"])[0];
  addTextThemeOverride(f, "z", "transformed", scheme.replace('<a:srgbClr val="123456"/>',
    '<a:srgbClr val="FFFFFF"><a:shade val="50000"/><a:alpha val="50000"/></a:srgbClr>'));
  const s = await open(f);
  try {
    assert.equal(style(s).color, "898989");
    assert.equal(style(s).colorOpacity, .25);
    const locations = records(s, "run.color")[0].resolvedFrom;
    assert.ok(locations.some(location => location.part === "/deck/transformed.xml" && /shade/.test(location.path)));
    assert.ok(locations.some(location => location.part === "/deck/z.xml" && /shade/.test(location.path)));
    const plan = await s.prepareSave({ mode: "no-op" });
    assert.deepEqual((await s.save(plan.plan)).bytes, new Uint8Array(encodeStructure(f)));
    assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked");
  } finally { s.dispose(); }
});

test("an explicit opaque color resets inherited alpha while unsupported color blocks RGB and alpha together", async () => {
  const f = fixture();
  f.parts["deck/layout.xml"] = f.parts["deck/layout.xml"].replace('<a:srgbClr val="445566"/>',
    '<a:srgbClr val="445566"><a:alpha val="25000"/></a:srgbClr>');
  f.set("deck/z.xml", "sld", inheritedShape({ text: textBody({ paragraphs: `<a:p><a:r><a:t>Inherited</a:t></a:r>${run(colorXml("srgbClr", "ABCDEF"))}${run(colorXml("srgbClr", "123456", '<a:hue val="120000"/>'))}</a:p>` }) }));
  const s = await open(f);
  try {
    const styles = s.document.slides[0].elements[0].text.content[0].runs.map(r => r.style);
    assert.deepEqual(styles.map(style => style.colorOpacity), [.25, 1, undefined]);
    assert.equal(styles[1].color, "ABCDEF");
    assert.equal(records(s, "run.colorOpacity")[2].status, "unsupported");
    assert.notEqual(styles[2].color, "445566");
  } finally { s.dispose(); }
});

test("dependency provenance has an independent exact budget and captured import options", async () => {
  const f = fixture({ color: colorXml("schemeClr", "accent1", '<a:lumMod val="50000"/><a:alpha val="50000"/>') });
  const baseline = await open(f);
  const count = baseline.report.textStyleInheritance.properties.reduce((total, p) => total + (p.resolvedFrom?.length ?? 0), 0);
  baseline.dispose();
  assert.equal(DEFAULT_SESSION_LIMITS.maxTextStyleDependencies, 1000000);
  assert.ok(count > 1);
  await assert.rejects(open(f, { sessionLimits: { maxTextStyleDependencies: count - 1 } }), error => error.code === "import-limit-exceeded" && /maxTextStyleDependencies/.test(error.message));
  const options = { projection: "content", sessionLimits: { maxTextStyleDependencies: count } };
  const pending = importPptx(encodeStructure(f), options);
  options.sessionLimits.maxTextStyleDependencies = 1;
  const s = await pending;
  s.dispose();
  for (const value of [0, -1, 1.5, NaN, Infinity, undefined]) await assert.rejects(open(f, { sessionLimits: { maxTextStyleDependencies: value } }), RangeError);
  const structure = await open(f, { projection: "structure", sessionLimits: { maxTextStyleDependencies: 1 } });
  structure.dispose();
});

test("shared long transform sources cannot amplify target provenance past its budget", async () => {
  const f = fixture();
  const transforms = '<a:lumMod val="100000"/>'.repeat(512);
  f.parts["deck/layout.xml"] = f.parts["deck/layout.xml"].replace('<a:srgbClr val="445566"/>',
    `<a:srgbClr val="445566">${transforms}</a:srgbClr>`);
  f.set("deck/z.xml", "sld", inheritedShape({ text: textBody({
    paragraphs: `<a:p>${Array.from({ length: 16 }, (_, i) => `<a:r><a:t>Shared ${i}</a:t></a:r>`).join("")}</a:p>`,
  }) }));
  await assert.rejects(open(f, { sessionLimits: { maxTextStyleProperties: 10000, maxTextStyleDependencies: 1024 } }),
    error => error.code === "import-limit-exceeded" && /maxTextStyleDependencies/.test(error.message));
  const s = await open(f, { sessionLimits: { maxTextStyleDependencies: 50000 } });
  try {
    const runs = s.document.slides[0].elements[0].text.content[0].runs;
    assert.equal(runs.length, 16);
    assert.ok(runs.every(run => run.style.color === "445566"));
  } finally { s.dispose(); }
});
