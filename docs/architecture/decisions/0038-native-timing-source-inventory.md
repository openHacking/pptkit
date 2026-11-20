# ADR0038: Read-only native timing source inventory

Status: accepted for source inspection only, 2026-10-26.

## Problem

Core foreground timing contracts need a reproducible source comparison before native effect mapping can be admitted. Raw timing trees include inactive MCE branches, foreign behaviors, subtargets and ambiguous identities. Flattening them into authored stages would grant unsupported semantics and risk modifying the wrong target.

## Decision

Expose opt-in `inspectPptxTiming(bytes, options?)` in Parser, using the existing OPC/XML/MCE context. Keep source-format identities and raw source locations here; add no Core dependency or event mapping. Default inspection/import reports stay compatible.

Publish flat, frozen observations with explicit effective/raw-only membership, separate part-local drawing/time identity indexes, lexical values, conservative candidate lookup and findings. Effective means XML membership only. Unknown/alias/subobject observations retain evidence without edit or execution authority. Global work/record/text/finding budgets reject completely; infrastructure budgets and cooperative abort remain active.

`nativePlayback` and `authoredMapping` explicitly remain `not-performed`. Native effect acceptance, source rewrites and actual playback require separate per-effect decisions. Public inspection is accepted; automatic mapping is deferred. Static preview, source preservation and headless reopen do not substitute for playback.

## Alternatives and consequences

Extending default import would add cost and accidentally imply mapping authority. A raw AST API would expose parser internals and complicate bounded transport. Keeping this only as a private research helper would prevent reproducible installed-consumer comparisons. The explicit plain-data API owns the inspection responsibility without new dependencies or packages.

Expanded paths can be large, so charge them before construction and keep independent XML limits. Output metrics describe captured data, not exact serialized bytes. Candidate identity is not stable across producer renumbering and must not be used as a persistent authored binding.

## Verification

See the [API contract and evidence](../../api/native-timing-inventory.md), eight Parser groups, two browser DOM parity groups, installed declaration/runtime consumers and optional Python ZIP/XML verification of two retained four-slide artifacts. No interactive native playback admission is granted. ADV-06.2.1.3 and ADV-06.1.2.3 remain open.

An additional [cross-package regression](../../../scripts/animation-mapping-parity.test.mjs) combines a foreign behavior and an inactive MCE branch containing a character-range target. A mixed unrelated/target text edit is blocked, restoring the target permits retry, and the saved unrelated edit retains the entire target slide and every other unchanged package part byte-for-byte. This fixture is a preservation guard, not a claim that arbitrary unknown timing can be interpreted or safely edited. It runs in `test:import`; full effect admission and the remaining combinations stay open.

A second fixture combines whole-object and character-range references to the same shape, a missing target ID and a foreign behavior. The inventory locates both references to the real owner and reports the missing ID. A mixed owner/bystander edit on the same slide is blocked; after restoring the owner, the bystander edit saves with exactly its expected character replacement and all other part bytes unchanged. This checks one explicitly constructed combination, not all possible opaque references or animation semantics.
