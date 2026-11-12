# Chart and authored-notes formatting origins

This extends [text-format origins](../api/text-format-origins.md) within Core normalization. There is no new IR field or source-preserving write capability. Current authoring has direct notes content; Parser's native notes-master/theme projection remains a separate domain.

## Verified cases

[Eight Core test groups](../../packages/core/test/chart-notes-origins.test.mjs) cover two masters interleaved across three slides plus a default theme domain, bar/line/pie chart defaults and overrides, concrete theme resolution, title-specific defaults, absent title/hidden legend, nested groups and shared layout charts, source null barriers, empty content/paragraph/run variants, re-resolution/reset and immutable snapshots. Every supported effective chart/notes leaf must have exactly one origin. Independently reading authored source paths must explain the effective value or explicit null; symbolic chart references are checked against the actual owning theme rather than compared as equal scalar values.

[Node/browser parity](../../scripts/multi-master-parity.test.mjs) compares full ordinary IR, immutable origins, Layout, SVG, warnings and PPTX bytes in Node, browser ESM and browser global with all network requests blocked. An independent DOM XML inspection checks actual chart title/axis fonts, sizes, colors and off flags, plus direct notes styles under independent slide master themes. An empty default notes body is correctly omitted by Exporter even though Core's logical empty notes formatting is reported. Pie charts do not invent axis XML or origin rows.

The [captured artifacts](chart-notes-origins-evidence.json) include the four-page source, complete IR/origins/Layout, four SVGs, Chromium preview and source/test dependency hashes. Eleven charts exercise chart text without adding unsupported chart families. The retained earlier two-page native source regenerates to SHA-256 `ef3d886dcf09d80e23b58cbbafb67cbb26be84fbb2a5c71797c2eb0a3219cd2d` exactly; its three-round native evidence is unchanged. No new native application run is claimed for the new fixture.

Installed Node and strict TypeScript consumers cover notes target narrowing, source clears and chart label properties while the original extendable target interface remains compatible. Public unions grow to include notes targets and chart/notes property paths; exhaustive consumers need to account for the new cases.

## Adversarial review and open requirements

The review checks five likely regressions: tracing changes ordinary IR; title defaults get attributed to the chart body; concrete chart theme values erase the symbolic declaration source; notes borrow slide placeholder formatting; and empty/reset content leaves stale origins. Focused tests cover each risk, including actual XML and source-path reads. Existing native/Parser reports are not treated as proof of new authoring capabilities.

Table regional/cell origins still require metadata carried through the actual merge chain, including implicit cells and compact merged-grid source positions. Full live-template/reset behavior, master artwork, independent authored notes master/theme styles, notes fields/layout authoring, general native chart typography and source style writes remain open.

## Check record

Build/typecheck and complete workspace `pnpm test` exited 0 (`.tmp/chart-notes-origins-{build,typecheck,full-test}.log`): Core190, Layout139, Parser696, Exporter59, Renderer20, CLI11, visual33, import121 and isolated installed Node/browser/strict NodeNext declarations. Eight focused Core groups and all four multi-master/browser parity groups pass. Lint, dependency boundaries, documentation links/contracts/typed examples and `git diff --check` pass (`.tmp/chart-notes-origins-lint.log`). Artifact verification matches 10 file hashes/byte counts, seven dependency hashes, 726 unique target/property records and exact retained native source bytes (`.tmp/chart-notes-origins-artifact-verification.log`). No commit or publication was made.
