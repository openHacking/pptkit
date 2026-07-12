import test from "node:test";
import assert from "node:assert/strict";

import { createPresentation } from "@pptkit/core";
import { resolveLayout } from "../dist/index.js";

test("resolveLayout reports the current slide count", () => {
  const presentation = createPresentation();
  presentation.addSlide();
  presentation.addSlide();

  const result = resolveLayout(presentation);

  assert.deepEqual(result, {
    slideCount: 2,
    status: "placeholder",
  });
});

