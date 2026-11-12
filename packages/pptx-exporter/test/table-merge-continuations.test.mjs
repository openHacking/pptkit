import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation } from "@pptkit/core";
import { generatePptx } from "../dist/index.js";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";

test("wide horizontal/vertical merges emit continuation flags at every original physical column", async () => {
  for (const span of [2, 3, 4, 7]) for (const leading of [0, 1]) {
    const document = createPresentation(), count = leading + span + 1;
    const rows = Array.from({ length: 4 }, (_, row) => ({ height: 20, cells: [
      ...Array.from({ length: leading }, () => ({ content: `Leading ${row}` })),
      ...(row === 0 ? [{ content: "Merged", colSpan: span, rowSpan: 3 }] : row === 3 ? Array.from({ length: span }, (_, column) => ({ content: `Final ${column}` })) : []),
      { content: `Trailing ${row}` },
    ] }));
    document.addSlide().addElement({ type: "table", columns: Array(count).fill(40), box: { x: 0, y: 0, width: count * 40, height: 80 }, rows });
    const xml = readGeneratedZip((await generatePptx(document)).bytes).get("ppt/slides/slide1.xml").toString();
    const physicalRows = [...xml.matchAll(/<a:tr\b[^>]*>([\s\S]*?)<\/a:tr>/g)].map(m => [...m[1].matchAll(/<a:tc(\s[^>]*)?>/g)].map(m => m[1] ?? ""));
    assert.equal(physicalRows.length, 4);
    for (let row = 0; row < 4; row++) {
      assert.equal(physicalRows[row].length, count);
      for (let column = 0; column < count; column++) {
        const inside = column >= leading && column < leading + span;
        assert.equal(physicalRows[row][column].includes('vMerge="1"'), inside && row > 0 && row < 3, `${span}/${leading}/${row}/${column}`);
        assert.equal(physicalRows[row][column].includes('hMerge="1"'), inside && column > leading && row < 3, `${span}/${leading}/${row}/${column}`);
        if (inside && row === 0) assert.ok(physicalRows[row][column].includes('rowSpan="3"'));
        if (inside && column === leading && row < 3) assert.ok(physicalRows[row][column].includes(`gridSpan="${span}"`));
      }
    }
  }
});


test("whole merged anchors retain four logical outer borders while physical merge interiors remain hidden", async () => {
  for (const span of [2, 3, 7]) {
    const document = createPresentation();
    const colors = { left: "AA00AA", right: "0000FF", top: "FF0000", bottom: "00AA00" };
    document.addSlide().addElement({ type: "table", box: { x: 0, y: 0, width: span * 40, height: span * 20 }, columns: Array(span).fill(40), rows: Array.from({ length: span }, (_, r) => ({ height: 20, cells: r ? [] : [{ content: "Merged", rowSpan: span, colSpan: span, style: { borders: Object.fromEntries(Object.entries(colors).map(([side, color]) => [side, { paint: { type: "solid", color }, width: 3 }])) } }] })) });
    const result = await generatePptx(document);
    assert.deepEqual(result.warnings, []);
    const xml = readGeneratedZip(result.bytes).get("ppt/slides/slide1.xml").toString();
    const cells = [...xml.matchAll(/<a:tc(?:\s[^>]*)?>([\s\S]*?)<\/a:tc>/g)].map(m => m[1]);
    for (const [side, tag] of [["left", "lnL"], ["right", "lnR"], ["top", "lnT"], ["bottom", "lnB"]]) assert.ok(cells[0].includes(`<a:${tag} w="38100"><a:solidFill><a:srgbClr val="${colors[side]}"`));
    // A non-anchor, non-perimeter slot must not acquire any visible border.
    if (span > 2) {
      const center = cells[span + 1];
      assert.equal((center.match(/<a:noFill\/>/g) ?? []).length, 4);
    }
  }
});
