# Notes source inventory verification — 2026-10-27

This delivery adds [ADR0056](../architecture/decisions/0056-notes-source-inventory.md) and the [explicit source-only API](../api/notes-source-inventory.md). It advances IMP-03.3.1/.2 without closing their complete inheritance/output/native authoring gates.

## Independent source checks

The [native-file runner](../../scripts/inspect-native-notes-inventory.mjs) hashes its four input evidence manifests and checks each actual PPTX against the previously recorded SHA-256 before inspection. It starts no native application. The [current manifest](notes-source-inventory-evidence.json) records 10 actual inputs, 98 ordered pages, 4 stored fields and 7,550 checked source references. They include the PowerPoint-authored control, A rect/copy third rounds, B and distinct per-notes override third rounds, three fresh author sizes and empty/typed insertion-style third rounds. These are fresh checks of captured files, not newly repeated GUI operations. The native control retains `datetimeFigureOut` cache `2026/10/26`, master `slidenum` cache `‹#›` and local number caches `1`/`2`; inspection on 2026-10-27 leaves all four literal values unchanged.

Each input receives a full JSON report snapshot and hash under a new `.tmp/notes-source-inventory/run-*` directory. The independent [DOM oracle](../../tests/support/notes-inventory-source-audit.mjs) uses raw element preorder, expanded names, literal attributes, actual text/CDATA, nonvisual ancestry, direct geometry components and cached field ancestry. It imports neither Parser nor Core. B/override page/master/theme/override identities additionally equal the existing captured independent domain maps. Browser global reports equal Node reports in full, without network requests. Each structure-session no-op returns exact original bytes, and its new inventory equals the original report.

`.tmp` inputs/snapshots are local ignored artifacts; a manifest alone does not restore them. Missing or mismatched files fail the runner. Original native versions/procedures remain in [A rect/copy](native-notes-pages.md), [B/override](notes-image-domains.md), [fresh sizes](authored-notes-pages.md) and [insertion styles](notes-insertion-styles.md). Prior failed fixtures and observations are retained.

## Focused and consumer checks

The [Parser tests](../../packages/pptx-parser/test/notes-inventory.test.mjs) cover both dialects; role/type matching independent of idx/name; groups, absent/ambiguous bodies and missing local text; master prompt isolation; root/inner MCE alternatives and ProcessContent; ambiguous inactive nonvisual structure without an invented identity; unknown islands; numeric identity aliases and prototype-shaped type strings; header/footer/nonvisual flags; picture/connector/frame/custom geometry; literal date/number/unknown caches; UTF-16 and empty caches; missing/actual unlisted masters; malformed/external edges; exact resource thresholds; option/byte capture; deep freezing; cancellation and retry. Large shared-master tests verify parts are scanned once instead of once per page.

The [browser parity test](../../scripts/notes-inventory-parity.test.mjs) compares complete Node/ESM/global reports for both dialects and wrapped/unwrapped nonvisual structures. Independent DOM checks validate raw sources, and a deliberately corrupted cache report is rejected. Limits, cancellation, invocation capture, immutability and zero network requests are tested. This is part of `test:import`.

The isolated packed-package suite exercises installed Node, browser ESM/global and strict NodeNext type consumers, without workspace path aliases. New public types and expected compile errors are covered by [consumer types](../../tests/consumers/types.ts); the [typed documentation example](../examples/notes-source-inventory.ts) is checked by the repository docs checker.

## Adversarial review and limits

1. Matching by idx or choosing a listed master could cross domains: exact type candidates and actual edges are asserted, missing edges remain missing, and distinct native override maps are independently compared.
2. MCE wrappers could hide valid bodies or count inactive bodies: effective ownership/nonvisual structure now uses the selected projection and maps back to raw source identities; raw-only branches remain separately visible. Review-discovered inner nonvisual/text wrappers and field-cache wrappers have regressions. The field-cache test first failed against the old reader; the corrected reader selects the ProcessContent and AlternateContent values with their original raw identities.
3. Cached dates/numbers or raw transforms could masquerade as evaluated output: fixed report markers, literal cache/geometry tests and public docs prohibit that interpretation.
4. Shared masters, duplicate IDs or malformed input could cause unbounded/ambiguous work: per-part caching, global output/candidate/findings budgets, explicit ambiguity and atomic failure are checked.
5. A new inspector could accidentally authorize writes or alter existing behavior: default inspection equality, exact native no-op bytes and unchanged Core/save profile contracts are checked. No production dependency was added.

This inventory does not resolve inherited geometry/cropping/visibility, evaluate fields, create notes artwork, render a full notes page, or authorize source style/structural writes. Full notes, general native font selection, shared-override repair admission and the remaining backlog stay open.


## Completed validation

`pnpm build`, final Parser/browser build, `pnpm typecheck`, lint/boundaries/docs examples and diff checks passed. The complete `pnpm test` command exited 0, including 117 import integration tests (`.tmp/notes-inventory-delivery-full-test.log`). The last inactive-MCE identity guard landed after that command's package phase; final Parser tests were separately rerun at 696/696 (`.tmp/notes-inventory-final-parser-test.log`), and final browser parity/native-file checks passed. The full command's later import phase used the final build. Earlier 692/694/695 package-phase counts are not substituted for the final Parser result. The field-wrapper regression's pre-fix failure remains in `.tmp/notes-inventory-field-wrapper-before.log`.

Final checks: `.tmp/notes-inventory-final-parity.log`, `.tmp/notes-inventory-native.log`, `.tmp/notes-inventory-final-typecheck.log`, `.tmp/notes-inventory-delivery-lint.log`. The existing dev-app large-chunk build warning remains; no performance admission follows from a successful build.
