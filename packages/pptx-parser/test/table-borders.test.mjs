import assert from "node:assert/strict";
import test from "node:test";
import { normalizePresentation } from "@pptkit/core";
import { importPptx } from "../dist/index.js";
import { cell, table, tableRow, tableBytes } from "./table-fixture.mjs";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";
const sides = { left: "lnL", right: "lnR", top: "lnT", bottom: "lnB" };
const line = (name, color, width = 1) => `<a:${name} w="${width * 12700}">${color ? `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` : "<a:noFill/>"}<a:prstDash val="solid"/></a:${name}>`;
const properties = values => `<a:tcPr>${Object.entries(sides).map(([side, tag]) => line(tag, values[side] ?? null, values[side] ? 2 : 0)).join("")}</a:tcPr>`;
const open = bytes => importPptx(bytes, { projection: "content", editSaveProfile: "text-runs-v3" });
const sourceTable = s => normalizePresentation(s.document).slides[0].elements[0];
function merged(options = {}) {
  const a = { left: "FF0000", top: "00FF00" }, b = { right: "0000FF", top: "00FF00" }, c = { left: "FF0000", bottom: "FFFF00" }, d = { right: options.right ?? "0000FF", bottom: "FFFF00" };
  const rows = [tableRow([cell({ text: "Anchor", attributes: 'gridSpan="2" rowSpan="2"', properties: properties(a) }), cell({ attributes: 'hMerge="1" rowSpan="2"', properties: options.legacy ? "<a:tcPr/>" : properties(b) })], 20), tableRow([cell({ attributes: 'vMerge="1" gridSpan="2"', properties: options.legacy ? "<a:tcPr/>" : properties(c) }), cell({ attributes: 'hMerge="1" vMerge="1"', properties: options.legacy ? "<a:tcPr/>" : properties(d) })], 30)];
  return table({ rows });
}

test("both dialects map independently declared edge colors/widths and three text saves retain all border XML exactly", async () => {
  for (const strict of [false, true]) {
    let bytes = tableBytes(table({ rows: [tableRow([cell({ text: "Start", properties: `<a:tcPr>${line("lnL", "112233", 1)}${line("lnR", "445566", 3)}</a:tcPr>` }), cell()], 50)] }), strict);
    for (let round = 0; round < 3; round++) {
      const session = await open(bytes);
      try {
        const style = sourceTable(session).rows[0].cells[0].style;
        assert.equal(style.stroke.paint.type, "none"); assert.equal(style.borders.left.width, 1); assert.equal(style.borders.right.width, 3);
        const slide = session.document.slides[0], next = structuredClone(slide.elements[0]); next.rows[0].cells[0].content[0].runs[0].text = `Round ${round}`; slide.replaceElement(next.id, next);
        const prepared = await session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready", JSON.stringify(prepared.report));
        const saved = await session.save(prepared.plan), before = readGeneratedZip(bytes), after = readGeneratedZip(saved.bytes);
        const beforeBorder = before.get("deck/z.xml").toString().match(/<a:tcPr>[\s\S]*?<\/a:tcPr>/g), afterBorder = after.get("deck/z.xml").toString().match(/<a:tcPr>[\s\S]*?<\/a:tcPr>/g);
        assert.deepEqual(afterBorder, beforeBorder); for (const [name, data] of before) if (name !== "deck/z.xml") assert.deepEqual(after.get(name), data);
        bytes = saved.bytes;
      } finally { session.dispose(); }
    }
  }
});

test("physical continuation outer edges reconstruct all four logical merge sides in both dialects", async () => {
  for (const strict of [false, true]) {
    const session = await open(tableBytes(merged(), strict));
    try {
      const style = sourceTable(session).rows[0].cells[0].style;
      assert.deepEqual(Object.fromEntries(Object.entries(style.borders).map(([side, stroke]) => [side, stroke.paint.color])), { left: "FF0000", right: "0000FF", top: "00FF00", bottom: "FFFF00" });
    } finally { session.dispose(); }
  }
});

test("unequal physical merge-side segments are explicitly source-only and survive allowed text edits", async () => {
  const bytes = tableBytes(merged({ right: "FF00FF" })), session = await open(bytes);
  try {
    assert.equal(sourceTable(session).rows[0].cells[0].style.borders.right, undefined);
    const finding = session.report.findings.find(f => f.operation === "import" && /physical right border segments/.test(f.reason));
    assert.ok(finding?.source.path); assert.equal(finding.tier, "unsupported");
    const slide = session.document.slides[0], edited = structuredClone(slide.elements[0]); edited.rows[0].cells[0].content[0].runs[0].text = "Edited"; slide.replaceElement(edited.id, edited);
    const prepared = await session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready");
    const xml = readGeneratedZip((await session.save(prepared.plan)).bytes).get("deck/z.xml").toString();
    assert.ok(xml.includes('val="FF00FF"')); assert.ok(xml.includes('val="0000FF"'));
  } finally { session.dispose(); }
});

test("legacy anchor-only merges retain prior logical projection, and direct edge mutations never gain source-save authority", async () => {
  const bytes = tableBytes(merged({ legacy: true })), session = await open(bytes);
  try {
    assert.equal(sourceTable(session).rows[0].cells[0].style.borders.left.paint.color, "FF0000");
    assert.ok(!session.report.findings.some(f => /physical .*border segments/.test(f.reason)));
    const slide = session.document.slides[0], original = slide.elements[0], edited = structuredClone(original); edited.rows[0].cells[0].style.borders.left.width = 9; slide.replaceElement(original.id, edited);
    const blocked = await session.prepareSave({ mode: "edit" }); assert.equal(blocked.status, "blocked"); assert.ok(blocked.report.diagnostics.some(d => d.code === "save-edit-unsupported"));
    slide.replaceElement(original.id, original); const noop = await session.prepareSave({ mode: "no-op" }); assert.equal(noop.status, "ready"); assert.deepEqual((await session.save(noop.plan)).bytes, Uint8Array.from(bytes));
  } finally { session.dispose(); }
});
