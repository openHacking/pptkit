# Testing Strategy

PPTKit tests each responsibility at its narrowest useful boundary and then verifies the complete authoring-to-PPTX path.

## Required workspace checks

Contributors run the same commands expected in CI:

```bash
pnpm build
pnpm typecheck
pnpm lint
pnpm test
```

`pnpm lint` includes repository structure, dependency-boundary, documentation-link, stale-contract, and typed documentation-example checks.

## Package tests

- **Core:** operation semantics, immutable snapshots, IDs, duplication/removal, full diagnostics, normalized defaults, inheritance, and detached IR.
- **Layout:** detached results, connector anchors/bounds, image contain/cover, and nested group traversal.
- **PPTX exporter:** deterministic ZIP structure, OOXML parts/relationships, editable feature serialization, runtime asset behavior, warnings, and Node output.
- **PPTX parser:** bounded ZIP/XML/OPC/MCE inspection, mapped-source provenance, whole-plan mutation rejection and snapshot lifecycle. Explicit text-run preservation profiles require independent payload checks, actual-candidate reprojection, source-part identity/range counterexamples, three-round mixed-deck tests, browser byte parity and installed consumers. Node file output injects each open/write/sync/readback/link/rename/cleanup failure, source and target races, cancellation boundaries, racing writers and committed warnings; real-file integration independently compares unaffected payloads. Optional native experiments separately record producer losses and unverified interactive/PowerPoint gates.
- **SVG renderer:** deterministic serialization, all current element families, input detachment, asset/fidelity warnings, escaping, and real-Chromium visual fixtures.
- **CLI:** stable command entry behavior appropriate to its current minimal surface.

Tests prefer focused documents that isolate one contract. Cross-package tests cover boundaries where a feature would otherwise be “modeled but not exported.”

## Export verification layers

1. Validate authoring and normalized structures.
2. Open generated ZIPs and parse every XML/relationship part.
3. Assert key XML semantics and package relationships.
4. Render named fixture decks and compare reviewable snapshots.
5. Open/save representative decks in PowerPoint and LibreOffice.

The shared [fixture corpus](../testing/fixture-corpus.md) automates XML parseability, relationship/content-type coverage, embedded workbook inspection, semantic assertions, and Chromium preview snapshots. These checks do not constitute full OOXML schema validation or native application rendering. The [application matrix](../testing/application-compatibility.md) tracks the remaining open/edit/save evidence separately. Performance/output-size measurements live in the [baseline guide](../performance/README.md).

SVG renderer fixtures run in pinned Playwright Chromium. Install the local browser once
with `pnpm exec playwright install chromium`; CI installs Chromium and its Linux
dependencies before `pnpm test`. Fixed viewport, device scale, color scheme, and a small
cross-platform pixel tolerance keep the preview snapshots reviewable.
The cross-package `test:import` command limits Node test-file concurrency to four
because many files start Chromium; all listed tests still run. This reduces simultaneous
browser resource pressure and does not change their assertions or native-application gates.
The standard `pnpm test` also checks that every root `*-parity.test.mjs` file is registered
in a reachable test suite. `test:table-style` includes table text direction; `test:composition`
includes unresolved image fits, parallelogram geometry and SVG-pair replacement/import.
Evidence and native-application experiments with separate runtime requirements remain opt-in;
registration does not establish native compatibility or complete test coverage.

## Fixtures and regressions

- Package fixtures live under `packages/<name>/test/fixtures/` when needed.
- Cross-package demonstrations live under `examples/` or typed documentation examples.
- Golden artifacts must be small, deterministic, and tied to a named behavior/regression.
- Large “everything” decks supplement rather than replace focused tests.
- Every fixed export regression adds the narrowest assertion that would have detected it.

## Review rule

A behavior change is incomplete without tests at the owning package and, when package boundaries are affected, an integration/export assertion. Public contract changes also update examples, API reference, and architecture documentation in the same change.

Page-transition tests verify immutable state and numeric validation, exact independent DOM attributes, three fresh output/import rounds, unchanged static SVG, source-only unknown/MCE branches and unchanged source timing during unrelated character saves. Actual native settings/playback remain a separate application gate.

`pnpm test:timing-research` runs the private finite animation time experiment (also in the root test chain). Its algebraic interval/seek evidence precedes public/native animation admission; it does not count as application playback coverage.

The timing research command also validates candidate target/event gating: wrong-page/object and early clicks, exact-time order, same-click consumption, stable snapshots and aggregate container budgets. This is independent of native/interactive playback certification.

The public finite-time utility adds Core tests with independently calculated intervals, 400 seeded relative/expanded-time oracles, exact overflow/zero/repeat boundaries, quotas and malformed transported-plan checks. Node/Chromium ESM/global and installed runtime/type consumers verify the public arithmetic contract. It does not add or certify native effect playback; no visual animation claim is made.

Public event-timeline tests add complete state tables for scope/readiness/equal-time consumption, repeated random-access seeks/new visits, deleted current targets, complete future-plan/log validation and shared timing/text quotas. An independent event-first state machine checks 300 generated programs/logs; 1,000 equal-time zero gates verify no event reuse. Browser/installed runtime and declaration consumers cover the public utility; native/controller dispatch remains unverified.
