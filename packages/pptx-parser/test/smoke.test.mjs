import assert from "node:assert/strict";
import test from "node:test";
import { zipSync, strToU8 } from "fflate";
import { inspectZipArchive, PackageInspectionError } from "../dist/index.js";

test("public archive inspection accepts an independently packaged archive and reports corruption", async () => {
  const bytes = zipSync({ "ppt/presentation.xml": strToU8('<presentation name="smoke"/>'), "ppt/media/image.bin": Uint8Array.of(0, 255, 1) });
  const inspection = await inspectZipArchive(bytes);
  assert.deepEqual(inspection.entries.map((entry) => entry.name), ["ppt/presentation.xml", "ppt/media/image.bin"]);
  assert.equal(inspection.expandedBytes, 31);
  await assert.rejects(inspectZipArchive(bytes.subarray(0, bytes.length - 1)), (error) => error instanceof PackageInspectionError && error.code === "archive-invalid");
});
