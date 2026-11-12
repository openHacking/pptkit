# Table formatting origin verification

This slice completes declaration tracing for currently supported logical cell formatting, including regional and unstyled tables. It does not complete AUT-05/AUT-07, native reusable style editing, source style writes or all geometry/content origins. [The API](../api/table-format-origins.md) distinguishes cell candidates from shared physical borders.

## Independent checks

[Eleven Core groups](../../packages/core/test/table-origins.test.mjs) check all supported effective style leaves exactly once, and independently read each authored path to compare the selected value or null barrier. They cover scalar/partial margins, presets with punctuation, regional priority, sparse and uniform strokes, atomic paint/bullet replacement, whole/per-slot script clears and reopens, empty paragraphs/runs, implicit cells, nested groups, interleaved master themes, reset/structural edits and immutable snapshots. Ninety-six independently generated rectangular partitions use a geometric region oracle and equal-valued paints/text sizes to require the correct declaration owner even when values cannot distinguish it.

The occurrence-path review found that a reference-to-index Map would lose earlier positions when one cell object was reused. The bounded grid now retains the actual authored index per anchor; an explicit aliased-object regression requires distinct paths for both occurrences. Implicit cells do not acquire a fictitious source index or local declaration.

[The public boundary test](../../scripts/table-origins-parity.test.mjs) compares full IR/origins/Layout/SVG/warnings/PPTX bytes across Node, browser ESM and global, with network blocked. Independent DOM reads verify actual fresh native cell margins, flags, paints, alpha and edges, along with XHTML content/padding. It asserts the fixture's cell candidate widths are all 0 while shared native right/bottom edges are 1/3pt after Layout conflict resolution. No source-to-visible-edge equivalence is invented. The command is included in `pnpm test:table-style` and complete `pnpm test`.

Installed Node consumers check sources on implicit regional cells; strict declarations accept cell text and border property paths. The original extendable inheritance target interface is unchanged.

## Captured evidence and limits

[Prepared evidence](table-origins-evidence.json) binds source, complete normalized/origin/Layout snapshots, three SVGs, Chromium preview, the captured fixture and source/test dependency hashes. [The preparation script](../../scripts/prepare-table-origins.mjs) creates a unique directory without replacing old artifacts. Existing regional/script materialization warnings remain visible.

The previously admitted two-page native source regenerates byte-exact at SHA-256 `ef3d886dcf09d80e23b58cbbafb67cbb26be84fbb2a5c71797c2eb0a3219cd2d`. The preceding chart/notes source also remains byte-exact, and its complete ordinary IR/origins match the retained snapshot. Those checks guard other scopes; they do not admit native behavior of this new table fixture. There is no new PowerPoint run, and the older table border/alpha/transform and native reusable-style limitations remain open.

## Adversarial review

Five risks are covered: inferred ownership after materialization; an atomic paint accidentally retaining lower alpha; clear barriers being lost or over-applied; object aliases/implicit cells fabricating source positions; and confusing a cell border candidate with the visible shared edge. The occurrence-index defect was corrected before the final validation run. Ordinary IR equality, full consumer/output parity and direct source-leaf reads support the metadata contract; native appearance is not inferred from them.

## Current check record

Final build/typecheck and complete `pnpm test` exit 0 (`.tmp/table-origins-final-{build,typecheck,full-test}.log`): Core201, Layout139, Parser696, Exporter59, Renderer20, CLI11, visual33, table-style5, import121 and isolated installed Node/browser/strict NodeNext declarations. This final command started after the occurrence-index correction and all production/fixture/test changes. The earlier full run is retained in `.tmp/table-origins-before-occurrence-full-test.log`. Lint, dependency boundaries, documentation links/contracts/typed examples and `git diff --check` pass (`.tmp/table-origins-final-lint.log`). Independent artifact checks match 11 hashes/byte counts, 12 dependencies, 2,138 unique target/property records and both retained source byte comparisons (`.tmp/table-origins-artifact-verification.log`). No commit or publication was made.
