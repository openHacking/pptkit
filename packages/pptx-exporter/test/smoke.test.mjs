import test from "node:test";
import assert from "node:assert/strict";

import { createPresentation } from "@pptkit/core";
import { exportPptx } from "../dist/index.js";

test("exportPptx returns a structured placeholder result", async () => {
  const presentation = createPresentation({ title: "Quarterly plan" });
  presentation.addSlide();

  const result = await exportPptx(presentation, {
    output: "./deck.pptx",
  });

  assert.deepEqual(result, {
    output: "./deck.pptx",
    slideCount: 1,
    status: "not-implemented",
  });
});

