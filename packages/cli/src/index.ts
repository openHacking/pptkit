import { HELP } from "./commands/arguments.js";
import { execute } from "./commands/execute.js";
export { CLI_REPORT_SCHEMA_VERSION } from "./commands/execute.js";
export interface CliRunResult {
  exitCode: number;
  /** Does not include the executable's final newline. */
  output: string;
  /** Errors use stderr; successful and comparison-difference reports use stdout. */
  channel?: "stdout" | "stderr";
}
export interface CliExecutionOptions {
  /** Capture relative input/report paths against this directory. Defaults to process.cwd(). */
  readonly cwd?: string;
  readonly signal?: AbortSignal;
}
/** Synchronous informational compatibility entry. File commands use runCliAsync. */
export function runCli(args: string[] = []): CliRunResult {
  if (!args.length || args.includes("--help")) return { exitCode: 0, output: HELP };
  return { exitCode: 2, output: "Use runCliAsync for file commands, or --help.", channel: "stderr" };
}
/** Local-file command orchestration over the public Parser APIs. */
export function runCliAsync(args: readonly string[] = [], options: CliExecutionOptions = {}): Promise<CliRunResult> {
  return execute(args, options);
}
