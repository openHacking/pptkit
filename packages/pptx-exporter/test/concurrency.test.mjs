import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation } from "@pptkit/core";
import { generatePptx } from "../dist/index.js";
import { createFixture } from "../../../tests/fixtures/presentations.mjs";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";

test("async image loading cannot share chart numbering across export invocations", async () => {
  const document = createFixture("charts", createPresentation);
  const expected = Buffer.from((await generatePptx(document)).bytes);
  for (const result of await Promise.all(Array.from({ length: 8 }, () => generatePptx(document)))) {
    assert.ok(Buffer.from(result.bytes).equals(expected), "each output matches its isolated export");
    const entries = readGeneratedZip(result.bytes);
    const chartNames = [...entries.keys()].filter((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name));
    assert.deepEqual(chartNames, ["ppt/charts/chart1.xml", "ppt/charts/chart2.xml", "ppt/charts/chart3.xml"]);
    for (const name of chartNames) assert.ok(entries.get("[Content_Types].xml").toString().includes(`PartName="/${name}"`));
  }
});
