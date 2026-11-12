# Source table styles and theme projection

Content import now exposes `report.tableStyleInheritance` with profile `drawingml-direct-regions-v1`. Text defaults below describe the legacy single-family path; explicit [script-families-v1](script-font-import.md) adds independent font slots under a separate `report.textFontProfile` marker. Parser selects a source table style, evaluates supported direct declarations in the target slide's theme/color-map scope, and materializes effective fills, borders and text defaults into ordinary Core cells. Core never receives source GUIDs, relationships or native table-style definitions. Two verified explicit bare references can separately map to semantic preset names; this does not resolve their appearance or override a source definition.

See the [typed example](../examples/table-style-import.ts), [ADR0043](../architecture/decisions/0043-source-table-style-projection.md), [native observations](../testing/table-style-import-evidence.json), [direct table import](table-import.md) and [authoring region styles](table-style.md). This is an experimental source-format profile. Supported solid/no-fill matrix references are described in the [matrix extension](#theme-format-matrix-extension). Full native cascade, other matrix paints/effects, property provenance for every metric, default text appearance, source style writes and interactive editing remain open.

## Style selection and source ownership

Parser recognizes Transitional and Strict expanded names, verifies the presentation's internal whole-part `tableStyles` relationship and `presentationml.tableStyles+xml` content type, and loads its effective `a:tblStyleLst` through the existing MCE processor. It indexes canonical brace-delimited GUIDs case-insensitively. Duplicate effective GUIDs are ambiguous and reject; inactive MCE alternatives do not become competing definitions. Unsupported definition identities are reported, not silently renamed.

An inline `a:tableStyle` under `tblPr` takes the inline path; a simultaneous `tableStyleId` is an invalid conflicting choice. Otherwise, an explicit style ID selects the registry entry, and an absent explicit reference can select the registry's `def`. A malformed explicit ID never falls through to that default. Missing definitions report `missing`; built-in style GUIDs are never inferred from an application-specific library. Invalid relationship multiplicity, external/fragment targets, wrong content types/roots and malformed interpreted booleans reject with existing typed package/import errors. Unchanged structural projection does not evaluate table styles.

Each table selection records source, canonical style ID where available, definition location/name, effective six booleans, and `inline/resolved/missing/unsupported`. A resolved identity establishes the selected definition, not full visual support.

## Effective direct-region profile

The profile recognizes `wholeTbl`, horizontal `band1H/band2H`, vertical `band1V/band2V`, first/last columns, first/last rows, and four corner regions. All six `tblPr` flags default to false and are interpreted independently of whether a region defines properties. A true first-row/column flag excludes the first physical line from band counting; a true last flag excludes an anchor starting on that final physical line. Bands alternate at physical anchor positions. Corners require the corresponding row and column flags.

The current bounded projection applies whole table, horizontal band, vertical band, first/last columns, first/last rows, then northwest/northeast/southwest/southeast corners, followed by explicit cell/body/paragraph/run declarations. Selected declarations merge by property, not by replacing an entire lower region. Source order is separate from Core's custom authoring order. Independent native evidence found that applying horizontal bands after vertical bands produced four wrong internal cell fills in the dual-band control; source projection now applies vertical bands later. The recorded named producer agrees on all fills in the three current controls. This does not certify all overlap cases or PowerPoint behavior.

Conditional-region projection is explicitly preview **fallback**, even for unmerged tables. Merged cells additionally use anchor-start region selection and full-rectangle boundary contact. Native merged-region precedence is not certified. The projection never duplicates merged text; physical continuations remain under the existing complete grid proof. Inconsistent physical border segments cannot become one logical edge: with inherited styles, an explicit unsupported noFill/zero candidate blocks the older inherited side rather than revealing it accidentally.

Supported declarations:

| Source | Projection |
| --- | --- |
| Region `tcStyle/fill` or local cell fill | Explicit noFill or solid RGB/system fallback/theme color, with supported ordered color/alpha transforms |
| `tcBdr` left/right/top/bottom/insideH/insideV containing a supported direct `ln` | Region exterior/interior sides mapped against region/cell bounds, then existing shared Layout conflict rules |
| Local physical `lnL/lnR/lnT/lnB` | Direct logical/perimeter projection, including existing uniform merged-side aggregation |
| `tcTxStyle` bold/italic `on/off/def` | Explicit true/false, or inherit for def |
| Explicit font collection or major/minor Latin `fontRef` | Materialized font name in the target slide theme scope; explicit collection alternatives must agree as described below |
| Color within `fontRef` | Default text RGB/alpha below the region's explicit color and later local declarations |
| Explicit table text color and alpha | Materialized RGB/alpha defaults below local list/paragraph/run declarations |

A new explicit paint replaces both color and alpha. Unknown color transforms, paint families, line semantics/arrowheads, unsupported font references, unsupported matrix entries and other unsupported explicit values block older declarations for that property and retain located diagnostics. They do not turn into a convincing older fill/font/line. Unknown sibling properties/extensions stay source-owned and audited. Standard cell/body margin and vertical alignment behavior remains in the existing table reader; this report does not yet trace every such metric.

The target slide's slide/layout/master theme overrides, verified master theme relationship and color mapping determine selected colors/fonts. Shared table-style definitions are cached by theme scope, so two masters can reuse the same style ID without sharing resolved RGB. An unsupported nearer theme component blocks an older component; no other master's palette is substituted. Explicit lt/ea/cs theme slots may resolve, but different script families still cannot project into this source profile's single family; script selection, shaping and application substitutions remain separate. Omitted source appearance still uses explicit projection defaults rather than an inferred native default, including missing table text color/font defaults. Native fill agreement alone does not prove text appearance.

## Theme format matrix extension

[ADR0044](../architecture/decisions/0044-table-theme-matrix-references.md) adds `tcStyle/fillRef` and `tcBdr/*/lnRef` without changing the report profile or introducing a new Core style identity. Format, color and font components select their nearest slide/layout/master override independently. A nearer present but incomplete `fmtScheme` blocks an older matrix. Matrix lists retain positional slots after MCE processing: an unsupported first element does not shift the second element into slot one.

| Reference index | Meaning |
| --- | --- |
| Fill 0 or 1000 | Explicit no fill |
| Fill 1–999 | `fillStyleLst[index - 1]`, if present |
| Fill 1001–4294967295 | `bgFillStyleLst[index - 1001]`, if present |
| Line 0 | No referenced line, projected as noFill with width zero |
| Line 1–4294967295 | `lnStyleLst[index - 1]`, if present; no background-list interpretation |

Indices are unsigned integers; malformed or duplicate interpreted fields reject. Bounds are checked against actual list length without allocating by index. Missing matrices, absent slots and unsupported selected entries report blockers rather than selecting another slot or uncovering a lower declaration. Solid/noFill paints and supported straight table lines are admitted. Neutral `cap=flat`, `cmpd=sng`, `algn=ctr` and miter limits do not alter straight edge geometry; compound/non-neutral lines, arrowheads and unsupported paint/effect families remain outside this subset.

The selected entry's `phClr` binds lazily to the reference color. A fixed RGB/theme entry does not need or evaluate an unused reference color. A required missing/unknown placeholder blocks the property. Palette, reference and placeholder modifiers compose in source order from the original RGB program, with a single final RGB quantization; alpha set/mod/off retain their ordered behavior. For example, black + reference tint 99500 + matrix shade 50000 yields `080808`, while rounding the reference first incorrectly yields `090909`. Reference alpha .4 followed by placeholder alphaMod .5 yields .2. The internal composition program never appears in Core or public reports.

The selected property's source is the reference `/@idx`; `resolvedFrom` includes the actual selected matrix and color/transform dependencies. Unsupported matrix audits retain theme-part locations. Matrix indexing, resolution and audit contribute to cumulative table-style work, and retained dependencies use the existing shared dependency limit. Supported scalar source text saves preserve original matrix/style payloads; this extension grants no style-writing permission.

Ten additional [Parser groups](../../packages/pptx-parser/test/theme-matrix.test.mjs), two [cross-package/browser groups](../../scripts/theme-matrix-parity.test.mjs) and installed consumers cover both dialects, index boundaries, placeholder composition, independent components, unsupported positional slots, shared masters, exact budgets, cancellation/retry and three source-save rounds. Independent DOM slot checks and a Python source oracle do not reuse Parser's matrix evaluator.

The [native evidence](../testing/theme-matrix.md) records five controls. Placeholder and no-line PDF fills agree; line widths remain partial. Background-list and transform controls disagree with the named LibreOffice producer, and the no-fill control fails to resave/export there. These are retained compatibility gaps, not adjusted expected values or full native acceptance.

## Reference text color and explicit script fonts

[ADR0045](../architecture/decisions/0045-table-text-appearance-and-uniform-fonts.md) extends selected text appearance without changing the report profile. Within one region, `fontRef` color is a default below a direct `tcTxStyle` color. Later selected regions and local list/paragraph/run declarations keep their existing precedence. New color replaces both RGB and alpha. An unsupported high-priority color blocks both older fields. Font eligibility is independent: a valid color can project even when the font index is unsupported. The font property's source now identifies `fontRef/@idx`; color provenance identifies the actual color declaration and theme/transform dependencies.

Explicit `font` collections use the shared Latin/ea/cs reader. Direct families or `+mj/+mn` lt/ea/cs slots may collapse to this source profile's single font only when their resolved families agree. Equal supplemental script entries retain dependencies too; differing, missing/empty or unresolved alternatives block an older family. Duplicate supplemental script identities and duplicate interpreted slots reject. This prevents silently ignoring a table's explicit alternate font. A nearer present font component still replaces that component as a unit; empty/missing slots do not borrow older typefaces.

This is not automatic language/script fallback or mixed-script font assignment. Theme supplemental selection, bidi/shaping, glyph coverage, font availability and application substitution remain separate gates. A plain table `fontRef` continues to project its Latin slot; it does not establish one-family equivalence for every font in the theme. Omitted font/color declarations still use projection defaults rather than inferred native appearance.

Twelve additional [Parser groups](../../packages/pptx-parser/test/table-text-appearance.test.mjs), two [browser/cross-package groups](../../scripts/table-text-appearance-parity.test.mjs), installed consumers and an [independent source/PDF experiment](../testing/table-text-appearance.md) cover this extension. The named LibreOffice producer disagrees on reference-only default color (white source PDF, black resaved PDF versus expected accent2); explicit region/local color controls agree. It substitutes LiberationSans for Arial. These observations leave native default appearance and complete script-font support open.

## Provenance and limits

`PptxTableStyleInheritance` contains `tables` and selected `properties`. A property has its cell's zero-based physical row/column, destination XML location, selected declaration source, mapped/unsupported status, value or reason, optional region, text level and theme/color-map/transform dependencies. Properties cover selected cell fills, logical border candidates, and selected table/local text fields. They describe projection choices, not a complete property cascade history or a source-mutation permission.

Paths count effective expanded names while retaining prefixes. Under MCE they describe the effective tree, not an arbitrary XPath into the unprocessed raw package. Raw text-save bindings retain their independent source-node identity proofs. A representative uniform physical border segment supplies the logical-side source; the report is not an exhaustive list of every continuation segment.

The following `sessionLimits` accumulate across all tables/parts in one import:

| Limit | Default | Meaning |
| --- | ---: | --- |
| `maxTableStyleProperties` | 500,000 | Selection records plus selected cell/text property records |
| `maxTableStyleDependencies` | 1,000,000 | Retained dependency references across those records |
| `maxTableStyleWork` | 2,000,000 | Table-style indexing/path/audit, declaration and per-cell region traversal units |

Limits are positive safe integers under the existing captured session-limit contract. Output records/dependencies are charged before publication; work is checked at table-style checkpoints. Theme parsing and the general ZIP/XML/OPC/MCE/content/report/save budgets retain their existing independent limits. These accounting units are not total wire bytes or hard peak-memory bounds. Abort/failure yields no usable partial session; retry requires no cleanup. Input bytes/options are captured before asynchronous work. Structure/no-op semantics and default source edit guards remain compatible.

## Saves, evidence and remaining gates

No-op saves retain exact original bytes. Explicit V3 table text saves can change admitted existing run characters while retaining original table-style/theme/relationship payloads and conditional flags; the source values are not regenerated from materialized Core styles. Style/metric/structural mutations remain blocked atomically. Restoring the baseline permits retry. Fresh output contains effective direct formatting and cannot recreate the original source style definition or flags; this is a different operation from preservation.

[Ten Parser groups](../../packages/pptx-parser/test/table-style-inheritance.test.mjs) cover two dialects, inline/reference/default/missing/ambiguous selection, MCE, flags/regions, local overrides, borders, false/alpha/font/theme maps/overrides, shared definitions under two masters, unknown blockers, three source-save rounds, exact session-wide limits, byte/option capture, cancellation/retry and heterogeneous merged-border blocking. [Two cross-package groups](../../scripts/table-style-import-parity.test.mjs) compare Node/ESM/global provenance and three complete source-save outputs in both dialects with no network; independent Chromium DOM resolves selected raw source fill paths in non-MCE fixtures. Fresh materialization and source style rejection/restoration are separate checks. Installed runtime/types consumers exercise the new report and limits.

The optional [independent runner](../../scripts/inspect-table-style-import.mjs) checks selected source fill locations/colors with Python ZIP/ElementTree and rejects deliberate source color corruption. Three valid generated controls exercise dual-band/corner regions, whole-table theme fill and local fill override. The named LibreOffice headless saved direct cell fills now agree for all 16 cells per control; original-source and resaved PDFs retain all control text, and region/theme PNGs were visually inspected. This evidence covers fill values and PDF text presence. It does **not** prove text color/typography, border/alpha/transform parity, retained reusable native style behavior, interactive editing or PowerPoint acceptance. Conditional previews remain fallback and the full TODO gates remain open.
