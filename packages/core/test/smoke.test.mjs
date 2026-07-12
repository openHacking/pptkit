import test from "node:test";
import assert from "node:assert/strict";

import { createPresentation } from "../dist/index.js";

test("createPresentation builds a minimal document shell", () => {
  const presentation = createPresentation({ title: "Roadmap" });
  const slide = presentation.addSlide();

  assert.equal(presentation.title, "Roadmap");
  assert.equal(presentation.slides.length, 1);
  assert.equal(slide.id, "slide-1");
});

