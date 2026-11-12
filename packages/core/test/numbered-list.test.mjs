import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, formatListMarker, validatePresentation } from "../dist/index.js";

test("list marker formatting is deterministic and validates the public input", () => {
  assert.equal(formatListMarker({ type: "none" }), ""); assert.equal(formatListMarker({ type: "bullet", character: "◆" }), "◆ ");
  for (const [n, expected] of [[1, "a"], [26, "z"], [27, "aa"], [28, "ab"], [52, "az"], [53, "ba"], [702, "zz"], [703, "aaa"], [32767, "avlg"]]) {
    assert.equal(formatListMarker({ type: "number", style: "alphaLowerPeriod", startAt: n }), `${expected}. `);
    assert.equal(formatListMarker({ type: "number", style: "alphaUpperPeriod", startAt: n }), `${expected.toUpperCase()}. `);
  }
  assert.equal(formatListMarker({ type: "number", style: "arabicParen", startAt: 53 }), "53) ");
  assert.equal(formatListMarker({ type: "number", style: "arabicPeriod", startAt: 53 }), "53. ");
  for (const n of [0, -1, 1.5, Infinity, Number.MAX_SAFE_INTEGER + 1]) assert.throws(() => formatListMarker({ type: "number", style: "arabicPeriod", startAt: n }), RangeError);
  assert.throws(() => formatListMarker({ type: "number", style: "unknown", startAt: 1 }), TypeError); assert.throws(() => formatListMarker({ type: "bullet", character: "" }), TypeError);
});

test("numbered-list starts use positive safe integers in ordinary text, shapes and notes", () => {
  for (const owner of ["text", "shape", "notes", "table"]) for (const startAt of [0, -1, 1.5, Number.MAX_SAFE_INTEGER + 1, Infinity]) {
    const deck = createPresentation(), slide = deck.addSlide(), content = [{ style: { bullet: { type: "number", startAt } }, runs: [{ text: "Item" }] }];
    if (owner === "notes") slide.setNotes(content);
    else if (owner === "shape") slide.addElement({ type: "shape", shape: "rect", box: { x: 0, y: 0, width: 200, height: 100 }, text: { content } });
    else if (owner === "table") slide.addElement({ type: "table", box: { x: 0, y: 0, width: 200, height: 100 }, columns: [200], rows: [{ height: 100, cells: [{ content }] }] });
    else slide.addElement({ type: "text", box: { x: 0, y: 0, width: 200, height: 100 }, content });
    assert.ok(validatePresentation(deck).some(d => d.code === "invalid-number-start"), `${owner}: ${startAt}`);
  }
});
