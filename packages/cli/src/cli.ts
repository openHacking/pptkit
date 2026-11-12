#!/usr/bin/env node
import { runCliAsync } from "./index.js";
const controller = new AbortController();
const cancel = () => controller.abort();
process.on("SIGINT", cancel);
try {
  const result = await runCliAsync(process.argv.slice(2), { signal: controller.signal });
  const destination = result.channel === "stderr" ? process.stderr : process.stdout;
  // A downstream consumer closing its pipe is not a parser failure.
  destination.on("error", (error: NodeJS.ErrnoException) => { if (error.code === "EPIPE") process.exitCode = 0; else { process.exitCode = 3; } });
  destination.write(`${result.output}\n`); process.exitCode = result.exitCode;
} finally { process.off("SIGINT", cancel); }
