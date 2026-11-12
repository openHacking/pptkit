# ADR 0064: Preservation-first reproduction and redaction

Status: accepted workflow design; automated minimization and sanitization are not implemented.

## Problem

A useful reproduction must still trigger the reported behavior, identify its exact input and environment, and avoid presenting a different generated file as the original. Presentation packages can contain shared resources, inactive compatibility branches, embedded files and application-specific references. A relationship graph alone cannot prove that removing a part is safe. Visible text and author fields are also only a subset of the places where sensitive information can occur.

## Decision

Keep the original immutable. Capture its byte hash, tool and runtime versions, operation, observed and expected results, named application/environment when relevant, and the relevant backlog IDs. Inspect explicit source roots with the Parser source-relationship closure, retaining findings, outside incoming owners and unreached identities. Treat coverage as scoped to relationship markup, including inactive branches, not as proof that a selected part set is a standalone document or that other parts are disposable. If inspection fails or reaches a limit, retain the error and stop that candidate; never treat a missing result as an empty dependency set.

Use a create-only whole-source no-op copy as the initial preservation candidate. Check its byte hash and run the same reproducing operation against it; retain the original if the observation does not reproduce. Exact equality verifies the copied bytes, not the application behavior. The [worked workflow](../../testing/minimal-reproduction.md) exercises this baseline with a synthetic authored fixture and the public closure runner. It is a reproducible starting point, not a smaller package.

Reduction, if needed, is a separately reviewed experiment: create a new candidate through an explicitly supported operation, retain its operation and source hash, inspect its package and dependencies, compare changed parts, and rerun the same failure predicate and any relevant native checks. Never infer equivalence from a successful save, a clean graph, a subset of parts or an empty capability report. Unsupported structural extraction remains unimplemented; reject ambiguous source ownership or unknown dependencies instead of silently dropping them. A failed candidate does not overwrite the source or replace the last verified candidate.

Treat redaction as an independent, scope-limited transformation, not a synonym for reduction. Source-bound text or metadata edits can change only their documented targets and can expose old values in reports/manifests. They do not remove secrets from notes, comments, hidden slides, properties, media, embedded packages, caches, unreachable parts, external targets, history or arbitrary opaque binary. No workflow in this decision certifies an input or report as sanitized. Before external sharing, the owner must independently review every artifact and decide whether its residual content is acceptable; when this cannot be established, reproduce with synthetic data instead.

Any future automated extractor must specify roots and control-part reconstruction, all-branch and incoming ownership handling, identity/references, conservative treatment of opaque parts, atomic create-only publication, resource limits, cancellation, failure reporting, and independent package plus reproducer validation. No such operation is admitted merely by this ADR.

## Evidence and scope

The [source closure contract](../../api/preservation-closure.md) and its cross-package tests include shared resources, inactive dependencies and unreached content. The [CLI regression](../../../packages/cli/test/commands.test.mjs) checks the documented authored-deck workflow, source/copy hashes, closure equivalence and failed root inspection without source mutation. These are automated package observations; no native application or general minimization has been verified. ECO-09.1.3 and its inherited gates remain open.
