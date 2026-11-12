import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation } from "@pptkit/core";
import { generatePptx } from "@pptkit/pptx-exporter";
import { importPptx } from "@pptkit/pptx-parser";
import { chartCandidates } from "../dist/commands/chart-candidates.js";
import { applyChartEdits } from "../dist/commands/chart-edits.js";

test("CLI v1 skips scatter coordinates without confusing category targets", async () => {
  const p = createPresentation(), slide = p.addSlide();
  slide.addElement({ type: "chart", chartType: "scatter", box: { x: 0, y: 0, width: 300, height: 200 }, series: [{ name: "S", xValues: [1, 2], yValues: [3, 4] }] });
  slide.addElement({ type: "chart", chartType: "bar", box: { x: 300, y: 0, width: 300, height: 200 }, categories: ["A"], series: [{ name: "B", values: [5] }] });
  const session = await importPptx((await generatePptx(p)).bytes, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
  try {
    const result = chartCandidates(session, "test");
    assert.equal(result.version, "pptkit-chart-edits-v1");
    assert.equal(result.unboundCount, 4);
    assert.equal(result.targets.length, 1);
    assert.equal(result.targets[0].expected, "5");
    const controller = new AbortController(); controller.abort();
    assert.throws(() => applyChartEdits(session, [{ ...result.targets[0], replacement: "9" }], controller.signal));
    assert.equal(session.document.slides[0].elements[1].series[0].values[0], 5);
    applyChartEdits(session, [{ ...result.targets[0], replacement: "9" }]);
    const plan = await session.prepareSave({ mode: "edit" });
    assert.equal(plan.status, "ready");
    const reopened = await importPptx((await session.save(plan.plan)).bytes, { projection: "content" });
    try { assert.equal(reopened.document.slides[0].elements[1].series[0].values[0], 9); } finally { reopened.dispose(); }
  } finally { session.dispose(); }
});
