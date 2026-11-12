# ADR 0032: Read-only CLI inspection and comparison

- Status: accepted experimental command/report contract
- Date: 2026-10-26

## Problem

The engine already exposes import capability/source reports and exact/conservative-XML package comparison, while the installed CLI still returns bootstrap text. Consumers need reproducible local inspection and comparison without implementing package parsing, obtaining private source state or mistaking cached/native observations for support.

## Decision

CLI depends on the public Parser package and owns argument parsing, regular-local-file capture, command orchestration and bounded human/JSON output. Parser retains all archive/XML/OPC/content and comparison rules. No new package, codec or external tool is introduced. CLI's TypeScript project references Parser; publish/install dependency ordering reflects this relationship.

Add `inspect` and `compare`, with explicit content/cache/text-save inspection and comparison profiles. All commands are read-only. Preserve the synchronous informational `runCli` entry; expose `runCliAsync` for file I/O and use it from the executable. Capture args/cwd/signal at invocation. Dispose imported sessions before returning reports. Errors become stable result/status codes and stderr; differences remain successful reports on stdout with exit 1. SIGINT forwards cooperative cancellation; downstream EPIPE does not print an internal stack.

The versioned envelope carries relative input paths, captured byte length/SHA-256, selected command/status, explicit native-verification absence and exact public Parser reports. It does not add editability because a finding is absent, synthesize task/evidence links, or certify native behavior. Additive fields are forward-compatible; incompatible envelopes require a schema version change. Existing nested Parser profiles retain their own scope.

Use bounded regular-file reads with pre-allocation size and post-read metadata/EOF guards. Nonblocking opens prevent FIFO hangs before rejection. CLI input/output bounds supplement, rather than replace, Parser expansion/work limits. Encode JSON incrementally and fail with a valid bounded error instead of silently cutting a report. Comparator difference truncation remains explicit and cannot imply equality. Human output includes located reasons and escapes terminal control/directional formatting characters.

## Verification and remaining work

Command tests compare exact public import/comparison results for Strict/Transitional owned fixtures, propagate SmartArt profiles, distinguish exact/XML equality and retained-difference truncation, reject malformed packages/arguments/nonregular files/limits, capture mutable options, process cancellation and check exact escaped UTF-8/newline boundaries. Actual Node processes verify exit codes and stdout/stderr; isolated npm-installed shims exercise inspect/compare and errors. Type consumers preserve synchronous help compatibility and check readonly execution options.

A complete operation/capability catalog with task/evidence references, streaming report transport, configurable deeper Parser budget profiles and broader authoring/save/tool commands remain separate work. ECO-09.1.1/.1.2 and parent tasks are not closed by this initial transport slice.
