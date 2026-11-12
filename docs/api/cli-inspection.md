# Read-only CLI inspection and comparison

`@pptkit/cli` provides local-file `inspect` and `compare` through the installed `pptkit` executable. These commands call public Parser APIs. They do not write source packages, fetch external resources or perform native application verification. The separate [create-only preservation save](cli-save.md) supports no-op and starting-number edits; [JSON generation](cli-generation.md) creates fresh files. Both use bounded input capture and Parser's Node file adapter; broader edits remain future work.

```sh
pptkit inspect deck.pptx --json
pptkit inspect deck.pptx --projection content --edit-save-profile text-runs-v3 --json
pptkit inspect deck.pptx --projection content --diagram-preview cache --diagram-text-comparison direct-paragraphs --json
pptkit compare before.pptx after.pptx --json
pptkit compare before.pptx after.pptx --profile opc-xml-v1 --max-differences 100 --json
```

Options may follow or precede file arguments after the command. `--` ends option parsing, so `pptkit inspect -- -deck.pptx` addresses a dash-prefixed filename. Duplicate, unknown, missing, invalid or cross-command options reject before opening files. Paths are local filesystem paths, not URLs or implicit stdin. Symlinks to regular files are supported; reports retain the normalized relative argument path. Directories, FIFOs, sockets and devices are rejected. Inputs are opened nonblocking, checked as regular files and bounded before allocation; descriptor metadata, EOF and path identity checks detect changes or path replacement observed during reading. These checks cannot exclude all concurrent mutations or changes after the final check. This is a captured-file inspection, not file-publication authorization or a filesystem-wide snapshot transaction.

## Reports and operations

JSON success/difference reports use `schemaVersion: "pptkit-cli-report-v1"`, `command`, `status: "ok" | "different"`, `inputs`, `nativeVerification: "not-performed"` and `report`. Each input has `path`, `byteLength` and SHA-256 of the captured bytes. Input paths are relative to the execution cwd and use `/`; package/source locations retain Parser logical paths. No timestamp is added. The same captured inputs/options produce deterministic output. JSON object property order has no semantic meaning; source arrays retain Parser order. Future additive fields must be ignored by consumers; an incompatible envelope change requires a new schema version. Nested Parser profiles keep their documented experimental boundaries.

Add `--package-summary` to `inspect --json` for an additional `packageSummary` with schema `pptkit-cli-package-summary-v1`. It reports archive input/expanded byte counts and entry/part counts, internal/external relationship counts, and a logical-part-sorted list of declared `image/*`, `audio/*` and `video/*` parts with compressed/expanded sizes and counts of incoming internal relationships. `mediaTotals` sums those unique parts, not references; shared parts are counted once. This option also works with `--capabilities`, leaving the chosen `report` unchanged. It uses the already inspected OPC graph and does not fetch external relationships. The classification is based only on the declared content type: it does not validate media bytes or include undeclared/non-media content types. Zero incoming relationships does not establish that deletion is safe, and relationship counts do not prove renderability, reachability or editability. ZIP entry sizes exclude container overhead. The full summary is subject to the same output budget; excess returns an error rather than a partial report. Default output is unchanged.

Add `--notes-inventory` to `inspect --json` for a `notesInventory` field containing the exact public [`inspectPptxNotes`](notes-source-inventory.md) result. It can accompany content inspection, capability views, candidate reports and `--package-summary` without changing their reports. It performs a separate bounded Parser inspection on the same captured source bytes, using the notes/XML/OPC budgets and the CLI input limit; it does not fetch resources. The combined JSON must fit `--max-output-bytes`, otherwise the command fails rather than returning a partial inventory. The inventory includes source text and stored field caches, which may be sensitive; it does not evaluate fields, resolve page geometry, render notes or authorize edits. Default output is unchanged.

With `--notes-inventory`, the positive-integer options `--max-notes-work`, `--max-notes-records`, `--max-notes-findings` and `--max-notes-text` set the respective Parser notes limits. Their defaults are 1,000,000 work units, 100,000 records, 10,000 findings and 1,000,000 UTF-16 code units. Each accepts 1–100,000,000 and is independent of `--max-output-bytes` and the underlying ZIP/XML/OPC budgets. Exceeding a notes limit returns `package-limit-exceeded` and exit 4 without a partial report. These counters are not total process-memory limits. Notes limit options without `--notes-inventory` reject before file access.

Without `--json`, the same flag appends a compact summary: notes size, finding counts, each slide's notes/master paths and body status, and per-part drawing, geometry and field record counts. These counts include raw-only observations where present; use JSON for the complete records, placeholder match statuses and stored field caches. Human output is subject to the same byte limit and terminal-control escaping as other reports; an over-limit result is an error, not a truncated summary.

`inspect` defaults to `--projection structure`. Its `report` is the public readonly `ImportReport` snapshot, including `inspection`, `findings`, identities and diagnostics. Each capability finding names feature, one of the seven import operations, tier, reason and source location. Missing findings are **not** evidence of support. This is a CLI transport of existing Parser findings, not a complete versioned capability catalog or task/evidence registry. Content mode is explicit. `--edit-save-profile` accepts the existing text-runs-v1/v2/v3 inspection/binding profiles only with content mode; it never prepares or writes an edited package. `--diagram-preview cache` also requires content mode. `--diagram-text-comparison direct-paragraphs` is independent; imported cache freshness remains unknown. Sessions are disposed before returning output.

`compare` defaults to exact `opc-payload-v1`; conservative `opc-xml-v1` normalization requires explicit selection. It returns the public `OpcPackageComparison`, with difference count, retained located differences and explicit truncation. All raw branches remain within Parser inspection. Comparison equality does not prove visual parity, semantic equivalence, native validity or source-edit safety. A truncated report still returns exit 1 whenever a difference exists; it cannot become equality because retained findings were exhausted.

Without `--json`, inspect prints counts and located feature/operation/tier/reason findings plus diagnostics. Compare prints the selected profile, equality/count/truncation and retained part differences. User-derived terminal control and directional formatting characters are escaped in human output. Success/equality and difference reports use stdout; errors use stderr. Reports include a final newline in the executable.

## Bounds, errors and cancellation

| Option | Default | Accepted range |
| --- | --- | --- |
| `--max-input-bytes N` | 67,108,864 per input | 1–536,870,912 |
| `--max-output-bytes N` | 1,048,576 including final newline | 512–67,108,864 |
| `--max-differences N` (compare only) | 1,000 | 1–1,000,000 |

Numbers require positive canonical decimal integers. Inputs are captured sequentially; comparing two files permits two per-file allocations. Parser archive/XML/OPC/import/comparison expanded-byte/work/report budgets remain in effect; increasing compressed input allowance does not remove them. No flags disable package validation or execute embedded actions.

Output encoding charges escaped UTF-8 tokens incrementally instead of materializing the complete JSON report first. A report exceeding its output budget returns error `output-limit-exceeded`; it does not emit partial JSON or silently omit report fields. Human output uses the same final-byte bound. This differs from the comparator's explicit retained-difference truncation. Underlying inspection reports still consume their own bounded Parser memory. Arbitrary consumer-provided values are not accepted for serialization.

Errors are versioned JSON envelopes with `status: "error"` and `error: { code, message, source? }`. Parser codes and logical source locations are retained. Filesystem errors use a bounded generic message without an OS stack or absolute filesystem path. There is no successful `report` on failure. If an underlying error itself exceeds the output bound, the output-limit error takes precedence. Command cancellation is cooperative across file reads and Parser checkpoints; `SIGINT` uses an AbortController, removes its handler when complete and returns 130 when processed by the command. Closing a downstream output pipe is handled without a stack trace.

| Exit | Meaning |
| --- | --- |
| 0 | Help, successful inspection, or equal comparison under its selected profile |
| 1 | Successful comparison found differences, or Core authoring validation found errors |
| 2 | Argument/API execution-option error |
| 3 | Input I/O, invalid or unsupported package/file |
| 4 | Input, Parser or output resource limit |
| 5 | Unexpected internal failure |
| 130 | Command cancellation |

## Programmatic use

<!-- doc-test: docs/examples/cli-inspection.ts -->
```ts
import { runCliAsync, type CliRunResult } from "@pptkit/cli";

export async function inspectLocalDeck(file: string, signal?: AbortSignal): Promise<CliRunResult> {
  return runCliAsync(["inspect", file, "--projection", "content", "--json"], signal ? { signal } : {});
}
```

`runCliAsync(args, { cwd?, signal? })` returns a `CliRunResult`; errors are results rather than rejected promises. Args/cwd/signal are captured before the first filesystem await. `output` excludes the executable's newline; optional `channel: "stderr"` identifies errors, otherwise use stdout. The synchronous `runCli` remains an informational compatibility entry for help; file commands use the async function. No import-state object or file handle is returned. `CLI_REPORT_SCHEMA_VERSION` exposes the envelope version literal.

Evidence: [command and encoder tests](../../packages/cli/test/commands.test.mjs), [isolated installed consumer/shim tests](../../scripts/installed-packages.test.mjs), readonly declaration consumers and [ADR0032](../architecture/decisions/0032-read-only-cli-inspection.md). Coverage includes both dialects, public-report parity, exact/XML comparison, changed and truncated outputs, malformed packages, FIFO rejection, dash/Unicode/control paths, option capture, cancellation, UTF-8 boundaries and separate executable streams/statuses. Native application validation and a complete task-linked capability catalog remain open under ECO-09/IMP-10.

For an end-to-end investigation using these commands and layout diagnostics, see the [debugging workflow](../testing/debugging-workflow.md).

The optional [versioned capability view](import-capabilities.md) is available with `inspect --capabilities`, exact `--capability-feature` / `--capability-part` filters and `--max-capability-subjects` (0–150000, default 1000). It adds `reportKind: "capabilities"` and returns the nested capability report instead of ImportReport only for this explicit option. Missing operations are unreported, tier disagreements are conflicting, and reference artifacts never certify current-input native/save validation.
