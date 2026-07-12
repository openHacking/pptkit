import test from "node:test";
import assert from "node:assert/strict";

import { runCli } from "../dist/index.js";

test("runCli exposes a help response", () => {
  const result = runCli(["--help"]);

  assert.equal(result.exitCode, 0);
  assert.match(result.output, /bootstrap/i);
});

