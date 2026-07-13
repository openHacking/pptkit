import test from "node:test";
import assert from "node:assert/strict";

import { assertBox, assertOptionalDimension } from "../dist/validation/geometry.js";
import { assertUniqueSlideIds } from "../dist/validation/collection.js";

test("geometry validation rejects non-finite and negative values", () => {
  assert.throws(
    () => assertBox({ x: Number.NaN, y: 0, width: 1, height: 1 }, "Box"),
    /non-finite x/,
  );
  assert.throws(
    () => assertBox({ x: 0, y: 0, width: -1, height: 1 }, "Box"),
    /negative width or height/,
  );
  assert.throws(() => assertOptionalDimension(-1, "Width"), /non-negative finite/);
});

test("collection validation rejects duplicate slide identifiers", () => {
  assert.throws(
    () => assertUniqueSlideIds([{ id: "intro", elements: [] }, { id: "intro", elements: [] }]),
    /Duplicate slide id "intro"/,
  );
});
