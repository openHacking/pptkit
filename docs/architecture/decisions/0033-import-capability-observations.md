# ADR 0033: Versioned import capability observation view

- Status: accepted experimental Parser/report/CLI contract
- Date: 2026-10-26

## Problem

Raw ImportReport findings expose feature/operation/tier/source/reason, but consumers must infer missing rows, group exact locations, handle contradictory observations and find related tasks/tests themselves. Absence must never become a support claim; static reference artifacts must not become current-input native validation.

## Decision

Parser owns a pure bounded `describeImportCapabilities` transformation and readonly types. It consumes only public projection/findings, constructs exact source/Core identity tuples, returns all seven operation rows, preserves duplicate observations and uses explicit reported/unreported/conflicting states. A tier exists only for agreeing observed tiers. No different feature/part/property transfers its capabilities to another subject. Native and candidate-save validation remain not-performed.

Known category/exact rules attach versioned task/artifact navigation references. They do not determine capabilities, authenticate input findings or prove that tests/native observations ran for this input. Unknown feature strings remain without fabricated references. This is a source-finding view, not a new Core model, full supported-feature catalog or runtime registry.

Keep schema and reference versions explicit. Ignore additive fields without executing unknown accessors; reject invalid known data/enums/locations and unknown quota names. Match filters exactly and validate/bound all source findings before selection. Use cumulative finding/text/work bounds, charge long-prefix sort comparisons and freeze detached output. MaxSubjects alone truncates successfully, with full matching counts/operation totals and complete retained subjects. The synchronous helper holds no external resources.

CLI exposes the view only under `inspect --capabilities`, with exact feature/part filters and retained-subject count. Its optional reportKind discriminator and nested version distinguish the selected report. Default inspect/compare output and all save permissions remain unchanged. Browser/root/global distribution and installed package types exercise the public contract.

## Evidence and remaining scope

Eight focused groups cover missing/conflicting/duplicate findings, every identity field and separator collisions, 40 seeded shuffled oracle passes, exact filters/whole-subject truncation, unknown fields/accessors/invalid enums, exact resource thresholds and retry, actual task/artifact references, and real both-dialect cached-preview/no-op reports. Browser integration independently reconstructs all public provenance fields and compares Node/ESM/global outputs. Its first oracle treated JSON object key order as provenance; it now compares each identity field explicitly. CLI/installed declarations and shims cover opt-in report selection and safe counts-only output.

Complete per-property source inventory, conditional candidate-save readiness, authored capabilities, broader exact task/evidence linkage and native mixed-document certification remain separate gates under IMP-04.4/ECO-09. Parent tasks are not closed by this view.
