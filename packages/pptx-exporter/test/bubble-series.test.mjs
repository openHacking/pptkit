import assert from "node:assert/strict";
import test from "node:test";
import { bubbleSeriesParts } from "../dist/ooxml/bubble-series.js";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";

const series = (name, xValues, yValues, sizes) => ({ name, xValues, yValues, sizes, color: "#4472C4" });

test("bubble series bind dense cache triples to separate worksheet ranges, including multi-letter columns", () => {
  const inputs = Array.from({ length: 10 }, (_, i) => series(i === 9 ? "A & <B>" : "same", [-0, i], [2, -3], [0.25, 16]));
  inputs[0] = series("short", [7], [8], [9]);
  const { xml, workbook } = bubbleSeriesParts(inputs);
  const worksheet = readGeneratedZip(workbook).get("xl/worksheets/sheet1.xml").toString();
  const cells = new Map([...worksheet.matchAll(/<c r="([A-Z]+\d+)"><v>([^<]+)<\/v><\/c>/g)].map(m => [m[1], m[2]]));
  assert.match(worksheet, /<dimension ref="A1:AD3"\/>/);
  assert.match(worksheet, /<c r="AC1" t="inlineStr"><is><t xml:space="preserve">A &amp; &lt;B&gt;<\/t>/);
  assert.equal(cells.size, 57);
  const nodes = [...xml.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/g)];
  assert.equal(nodes.length, 10);
  for (const [i, [, node]] of nodes.entries()) {
    assert.match(node, new RegExp(`<c:idx val="${i}"/><c:order val="${i}"/>`));
    assert.match(node, /<a:srgbClr val="4472C4"/);
    assert.match(node, /<c:bubble3D val="0"\/>$/);
    for (const [tag, values] of [["xVal", inputs[i].xValues], ["yVal", inputs[i].yValues], ["bubbleSize", inputs[i].sizes]]) {
      const data = node.split(`<c:${tag}>`)[1].split(`</c:${tag}>`)[0];
      const [, from, to, end] = data.match(/<c:f>Sheet1!\$([A-Z]+)\$2:\$([A-Z]+)\$(\d+)<\/c:f>/);
      assert.equal(from, to);
      assert.equal(Number(end), values.length + 1);
      assert.match(data, new RegExp(`<c:ptCount val="${values.length}"/>`));
      const cache = [...data.matchAll(/<c:pt idx="(\d+)"><c:v>([^<]+)<\/c:v><\/c:pt>/g)];
      assert.equal(cache.length, values.length);
      for (const [j, [, index, value]] of cache.entries()) {
        assert.equal(Number(index), j);
        assert.equal(Number(value), values[j]);
        assert.equal(cells.get(`${from}${j + 2}`), value);
      }
    }
  }
  assert.match(nodes[9][1], /Sheet1!\$AB\$2:\$AB\$3/);
  assert.match(nodes[9][1], /Sheet1!\$AC\$1/);
  assert.match(nodes[9][1], /Sheet1!\$AD\$2:\$AD\$3/);
  assert.equal(cells.has("A3"), false);
});

test("bubble series reject blank coordinates, nonpositive sizes and worksheet overflow", () => {
  for (const item of [series("", [1], [2], [3]), series("empty", [], [], []), series("length", [1], [2], [3, 4]),
    series("blank", [null], [2], [3]), series("hole", Array(1), [2], [3]), series("infinite", [1], [Infinity], [3]),
    series("negative", [1], [2], [-3]), series("zero", [1], [2], [0]), series("row limit", Array(1_048_576), [], [])]) {
    assert.throws(() => bubbleSeriesParts([item]), RangeError);
  }
  assert.throws(() => bubbleSeriesParts(Array(5462)), /5461/);
});
