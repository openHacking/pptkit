# Import and Export Boundaries

Exporting a PPTX and parsing a PPTX are related operations, but they are not inverse operations.

This boundary matters for both architecture and user expectations.

## Why They Are Not Inverses

PPTX packages contain behavior and structure that do not map cleanly back into a single normalized document model.

Examples include:

- master and layout inheritance
- theme mappings
- placeholders
- grouped coordinate transforms
- embedded media relationships
- compatibility extensions
- unknown or vendor-specific OOXML

Because of this, parsing is not just “undoing export.”

## Export Responsibilities

`@pptkit/pptx-exporter` should:

- accept normalized document state
- produce valid package structures
- enforce export-time constraints
- fail early when required invariants are not met

The exporter supports rich text, images, shapes, connectors, groups, tables, themes, layouts, placeholders, and notes. Recoverable image loading failures are represented as warnings and the affected image is omitted; Core validation failures remain hard errors containing all diagnostics.

The experimental content projection resolves supported same-scope straight connector references after parsing the complete slide forest. Core stores authored target IDs/anchors; Parser owns OOXML IDs, route decoding and diagnostics; Layout resolves geometry; Exporter maps the identities into part-local native drawing IDs only when `connectorMode: "native"` is explicitly selected. Unsupported attachments use cached free endpoint coordinates with an unmapped finding, not a live attachment. Direct safe click actions map to Core action values; Parser does not fetch URLs or execute source actions. Internal links resolve against the effective slide relationship graph.

The content projection also recognizes a narrow custom linear-path profile emitted for Core's round connector path modes: one literal no-fill path using move/line/close commands, bounded path coordinates, a solid sRGB line with optional literal alpha, round cap and join, and solid dash. Rotation and flips are baked into connector vertices around the source frame center; group-local coordinates remain local. Other custom geometry and line styling remain source-only with a finding. This is semantic projection, not a source edit binding; no-op save still preserves source bytes. Three fresh projection/export rounds cover open and closed samples and one transparent line, four rotation/reflection samples have Chromium vertex comparisons, one group-local sample is checked, and one PowerPoint-moved sample has been manually reimported.

Parser maps native section member names onto Core slide labels using validated source numeric identities. An explicit collection projection also maps section names and membership to independent Core identities; source GUIDs, original member order and unknown extensions remain in the Parser inspection/report. Core's authored collection fresh-exports as native sections subject to explicit membership constraints, while the legacy Core label remains PPTKit slide metadata only. Preservation-aware structural reconciliation remains separate work. See [ADR 0067](decisions/0067-native-section-authoring.md).

It should not assume that every PPTX in the wild could have been produced by the same normalized model.

## Parse Responsibilities

`@pptkit/pptx-parser` should:

- inspect package parts and relationships
- normalize understood structures into PPTKit-readable forms
- retain diagnostics about unsupported or partial results
- preserve unknown structures where possible

It should not assume that every parsed construct can be turned into a clean, fully editable canonical structure.

## Capability Tiers

Parser behavior should be thought of in capability tiers rather than a binary success/failure model.

Useful tiers include:

- fully normalized and editable
- normalized with partial degradation
- preserved but not fully understood
- unsupported and explicitly rejected

This framing helps the parser remain useful before full parity exists.

## Preservation-Oriented Thinking

When the parser encounters unsupported OOXML, silent data loss should be treated as a last resort.

Where practical, the system should preserve unknown structures or package fragments so that:

- round-trip damage is reduced
- diagnostics stay actionable
- future parser improvements can revisit preserved data

## Round-Trip Stability

Round-trip stability is an important long-term goal area, but it should not be documented as a blanket guarantee for every PPTX feature.

Instead, the architecture should assume:

- some structures round-trip cleanly
- some structures normalize with degradation
- some structures are preserved without full semantic understanding

That is a more honest foundation for future implementation work.

## Independent source comparison

Parser's [OPC comparator](../api/package-comparison.md) validates and compares original source packages without importing Core or calling Exporter. Exact payload comparison is the default; conservative XML normalization is explicit. Full document retention includes prolog/epilog, namespace bindings and all MCE branches, so effective projections cannot hide source loss. Comparison budget, cancellation and report truncation are independent of per-input inspection budgets. See [ADR 0005](decisions/0005-opc-source-comparison.md). Effective fresh-output semantics, renaming bijections, native validation and edit-aware preservation remain separate gates.

## Source relationship closure

Parser's [closure inspector](../api/preservation-closure.md) uses an internal relationship-only tree-retention mode to analyze all raw declarations, including inactive MCE branches, without retaining ordinary XML trees. It reports selected parts/footprint, outside incoming owners, unreached parts and located incompleteness/identity findings. Metadata is detached and frozen; no source trees, cleanup authority or edit-save capability enter Core or consumers. [ADR 0006](decisions/0006-source-relationship-closure.md) separates this graph analysis from arbitrary XML/binary references and later preservation-state reconciliation.

The private [XML character-edit primitive](decisions/0007-xml-character-edit-boundaries.md) binds lexical ranges to the complete validated source preorder and admits only explicitly identified character-only elements. It retains source encoding and copies unselected lexical regions verbatim, then revalidates candidate XML. It has no public export, Core run mapping, OPC writer or edit-save permission. The separate [text-run session profile](decisions/0008-direct-text-preservation-save.md) supplies bounded source ownership and package/session orchestration.

## Accepted import/save ownership decision

[ADR 0001](decisions/0001-import-preservation-save.md) assigns imported source state and preservation-save orchestration to Parser, with a future one-way public dependency on Exporter for explicit semantic rebuild. Core/IR stay format-independent. The decision defines no-op/edit/rebuild modes, prepared snapshot-bound plans, exact loss consent, disposal/cancellation and runtime conformance gates. The ADR alone is not runtime evidence. [Import sessions](../api/import-session.md) implement the initial lifecycle slice and an explicit [existing text-run save profile](../api/text-edit-save.md); complete semantic import, structural reconciliation and rebuild remain pending.


## Bounded XML substrate

Parser owns [per-part XML inspection](../api/xml-inspection.md), separate from ZIP validation and future OPC graph interpretation. The private adapter delegates XML grammar/namespace checks to pinned saxes 6, wraps fatal decoding and input/node/depth/content/pending-span budgets, rejects DTDs and never resolves external entities. This avoids duplicating a general XML grammar inside format-specific import code. The upstream library is archived; independent browser comparisons and wrapper regression tests are required evidence for the selected profile.

A private optional ordered tree supports future content-type/relationship readers, but the public API returns inventory only. Neither this projection nor a well-formed XML result establishes lossless source preservation or valid OOXML. Source bytes and package ownership stay with Parser sessions; envelope inspection alone does not create a session. Core, Layout and Exporter acquire no XML parser dependency.


## OPC graph ownership

[OPC inspection](../api/opc-inspection.md) composes private ZIP payload reading, namespace-aware XML and format-specific logical-name/content-type/relationship checks in Parser. Public inventory has no mutable trees or payload maps. External targets stay unresolved strings; shared/cyclic internal relationships are legal graph edges, not recursive loading instructions. Shared XML budgets are enforced by a private reader observer during parsing. This adds no runtime dependency or Core/Layout/Exporter import and does not make Exporter and Parser inverses.

## Presentation entry identification

Parser owns [PresentationML entry inspection](../api/presentation-inspection.md) above the OPC layer. It consumes one detached inventory, resolves the root officeDocument relationship and checks the main type/root dialect. It does not add private trees, source handles or OOXML to Core. Classification is distinct from schema validation, MCE processing and semantic import; unknown descendant markup remains uninterpreted.

## MCE processing ownership

Parser owns the [MCE processor](../api/markup-compatibility.md) between bounded XML reading and format-specific interpretation. Consumer namespace capabilities and application-defined extension islands are explicit configuration. Iterative processing keeps source trees intact, reports source element identities and returns a private expanded-name projection; it is not an XML serializer or preservation store. OPC relationships use a fixed configuration and expose decisions alongside effective edges. Source XML budgets include every branch, while relationship budgets count effective edges. Core/Layout/Exporter gain no MCE dependency; no capability is inferred from input namespace declarations.

## Structural presentation reader

[inspectPptx](../api/pptx-inspection.md) adds source ordering, exact EMU sizes and typed structural references. A private optional-retention OPC source reader produces bounded trees in the original parse pass; public envelope APIs still return inventory only. Structural readers index relationships by source and identity, cache effective MCE roots, and report unresolved inheritance without fabricating defaults. The reader does not create a Core document or attach preservation state. Trees and payloads remain private to Parser; no dependency direction changes.

## Document metadata reader

The structural Parser reader now uses the same private OPC XML source to read package-linked core, extended and custom properties. Format-specific values stay in Parser inspection types, with source part/relationship identities, lexical scalar strings and explicit uninterpreted descriptors. A bounded subset of inspected custom properties projects into Core/IR; unsupported values remain inspection-only. Core also supports authored typed custom properties for fresh export. Neither path represents arbitrary source variants or source identities or implies a preservation handle. Core has its own OPC markup rules; extended/custom property MCE uses a separate namespace profile. Calendar validation and complex variant interpretation remain outside this inspection profile. Long metadata lists use the captured cancellation checkpoint; no extra decompression, XML pass, dependency or network lookup is introduced.

## Native section inspection

Parser interprets the named MS-PPTX section extension after effective presentation ordering is resolved. Source GUIDs, names and numeric slide references remain format-owned; they are not equated with Core's per-slide section labels. Unknown presentation/section extensions expose shallow descriptors, never raw payloads or preservation handles. Section membership lookup uses a slide-ID map; traversal is cooperative and shares the original XML budgets.

The private MCE processor accepts a format-reader extension predicate without changing its public options. The structural reader activates it only for the main part's known section URI under the presentation → extLst ownership path. Application-parent tracking skips MCE/ProcessContent wrappers while retaining the raw parent for syntax checks. Opaque ancestors and unknown nested islands retain precedence. This permits inherited MCE scopes to work inside recognized section content without reparsing a detached subtree or treating every extension as understood.

## Structural import session runtime

`importPptx` requires explicit structure or content projection. Parser uses only Core public exports to construct a valid dimensions/metadata/slide-order projection and normalize detached snapshots. Source numeric IDs and OOXML provenance remain in frozen Parser reports. Full source bytes stay private; whole-package byte retention avoids traversing relationship cycles and includes unreferenced parts, but provides no general edit-save closure guarantee. No Exporter dependency is added until semantic rebuild can enumerate and verify its losses.

Plans are closure-owned records tracked by a per-session WeakMap and an iterable active-reservation set. They are not attached to arbitrary Core documents or a global registry. Frozen public plan/session records expose no candidate bytes or mutable private graph. Output reservations are acquired before allocation, count in-flight work, and release on failure/disposal; full normalized equality gates no-op and plan freshness. Default edit and all rebuild requests return explicit blocked reports with no assurance. The selected text profile can prepare a loss-free scoped-preservation plan. Serialization fails and source baselines never advance implicitly.

### Direct text-run preservation save

`text-runs-v1` is an explicit content-import option. The private MCE projector retains effective-to-raw origins; Parser binds each mapped text-box run to raw element preorder, expanded name and baseline characters. Breaks, mixed source atoms and synthetic runs remain unmapped. These identities live in Parser reports/private state, never Core or IR.

The separate [V2 owner profile](../api/text-owner-save.md) adds mapped shape text and direct notes bodies. Drawing bindings retain real Core element IDs; notes provenance exposes slide/notes-part identity without fabricating an element. Raw source guards include notes parts and scope targets by part plus drawing identity. Both owner families share report/analysis budgets and the existing whole-candidate comparison/reprojection pipeline; notes-master/inherited prompts and unmapped families remain unbound. [ADR 0009](decisions/0009-text-owner-preservation-save.md) records compatibility and ownership.

A pure normalized-snapshot analyzer classifies the entire candidate and permits only bound scalar run changes with fixed topology. Unsupported changes block the complete plan. Shared raw atoms must receive compatible proposals. Separate raw-source guards include inactive branches, known text/sub-object targets and signature/VBA declarations; incomplete source relationship coverage blocks.

The explicit [V3 table profile](../api/table-text-save.md) collects anchor runs only after full physical-grid validation. Public table provenance separates row/compact Core cell index from physical source column, retains the owning table element ID and does not invent cell elements. Scalar identity includes row/cell coordinates. Unmapped tables publish no bindings; missing bodies/synthetic runs remain unbound. The analyzer indexes cell paths only for table owners that have bindings, preserving the earlier profiles' preprocessing scope. [ADR 0010](decisions/0010-table-text-preservation-save.md) records the boundary.

Parser applies lexical XML changes, then a private ZIP32 writer preserves untouched compressed local records and central metadata while rebasing offsets. It retains every source part and performs no resource deletion or identity rewrite. Independent exact OPC comparison restricts differences to approved XML payloads. Reprojection of actual candidate bytes must equal the entire planned Core snapshot before the final freshness check. These modules add no Exporter dependency.

Edit preparation reserves the full per-plan output budget and shrinks it when ready. Per-operation controllers propagate caller abort/disposal into child parsing/writing/validation; a per-session relay map uses one listener per caller signal. Completion removes the relay, and disposal detaches caller listeners synchronously while children unwind. Failure releases reservations and returns no partial candidate. See [ADR 0008](decisions/0008-direct-text-preservation-save.md) for the accepted boundary and [the profile contract](../api/text-edit-save.md) for evidence and limitations.

### Node file publication

The Parser root stays filesystem-neutral while the `@pptkit/pptx-parser/node` entry owns local path policy. It canonicalizes a stable source parent, captures a no-follow regular file in bounded chunks and returns a process-local authentic source receipt alongside the normal session. Output uses a flushed temporary inode in the destination's canonical parent. Create publishes with an exclusive hard link; authenticated replace rechecks path, inode/metadata, SHA256, parent and temporary identities and then atomically renames same-directory.

The runtime-neutral synchronous `assertSavePlanCurrent` shares the private save ownership/liveness/consent/freshness implementation without copying bytes. Node calls it immediately before link/rename dispatch, closing the JavaScript await gap after asynchronous temporary work. Postpublication cleanup/sync failures are committed warnings; prepublication failures preserve the original and retain a retryable plan. Standard portable Node cannot compare-and-swap between the final check and rename, so external final-interval writers are outside the guarantee. Core/Layout/Exporter receive no filesystem or Node dependency. See [ADR 0011](decisions/0011-node-file-publication.md) and [the public contract](../api/file-save.md).


## Partial content projection

The explicit [content import profile](../api/content-import.md) reuses the structural reader's private context and retained/effective trees from one bounded OPC pass. Parser-owned readers map direct text, seven preset shapes, embedded raster pictures and nested groups into public Core inputs. A node/attribute consumption audit emits bounded, located unsupported-property findings; object IDs and effective source paths stay in Parser reports. No raw XML, source handles or dependency enters Core/Layout/Exporter.

The notes reader reuses that same rich-text interpretation and source-property audit, selecting a direct body placeholder and mapping only its paragraphs/runs/clicks into Core notes through `setNotes`. Notes identities refer to the owning Core slide, not fabricated drawing elements. Body-frame geometry remains unconsumed so it cannot be mistaken for imported notes-page layout. Notes master/theme inheritance and other page artwork remain source-only; Exporter uses its existing notes-page defaults for fresh output.

The [placeholder dimension resolver](../api/placeholder-inheritance.md) separately indexes effective slide/layout/master declarations and matches source index/type within verified part relationships. Per-property EMUs and source paths stay in the Parser report; Core receives local boxes and bounded preset names. Parent part caches and a declaration budget prevent repeated scans and ambiguity-list amplification. Master/layout prompt text is not copied. This does not replace the Core multi-master authoring contract or implement full theme/style inheritance, dynamic rebinding or preservation edits; see [ADR 0003](decisions/0003-placeholder-dimension-projection.md).

Only content import retains the validated archive payload map through projection. Parser resolves picture relationships and reads bounded PNG/GIF/JPEG header dimensions, then registers immutable bytes through the public Core asset API. Canonical source parts own asset identity; per-part and aggregate candidate byte budgets precede registration. Original source state and relationship provenance stay in Parser. External-only pictures do not become Core URL assets, and neither import nor later preview/export can implicitly load them. Header parsing is separate from full image decoding and does not claim EXIF/color/animation fidelity.

Direct styles and neutral/default projection values are distinguished from unresolved inheritance. Group child origins are rebased in Parser before Core/Layout consume local coordinates. The same session baseline/freshness/no-op rules apply to content projections. Source reconciliation beyond the selected text-run profile and lossy edit/rebuild remain blocked; direct fresh export by a consumer is explicitly separate from preservation save.

### Text color projection boundary

Text color projection uses a Parser-private ordered transform evaluator. Palette/map/transform paths remain in Parser provenance; Core/IR carry only RGB and optional semantic `colorOpacity`. Layout does not change text geometry based on alpha; Renderer and Exporter use the same semantic value. Property and dependency budgets separately bound shared-source provenance. See [the color profile](../api/text-color-transforms.md) for precision, blockers and recorded native differences.

### Table projection boundary

Parser's private table reader validates source EMU grids and rectangular merges before interpreting anchor text/actions. It converts physical cells into compact Core anchors with O(columns) active occupancy; optional absent bodies become empty content. Nonpositive dimensions, sparse/excess rows or nonempty continuations remain source-only, while malformed mapped values/invalid merge references fail the session. Row/column counts and traversal share the original XML/report/cancellation budgets.

Frame and grid dimensions stay independent. Mismatch is a located preview fallback rather than silently synthesized sizing. The existing rich-text/action reader interprets anchor content; a bounded cell-style reader maps direct fills, supported sparse per-edge strokes, uniform merged-perimeter segments and nonnegative margins/anchoring. Heterogeneous/partly missing merged segments, unsupported table styles, other continuation properties and inheritance remain source-owned and audited. No OOXML types or cell source IDs enter Core; cells retain table-owned source paths in Parser reports. Exporter puts table rotation/flips on a native group with authored root identity/action and a private local editable table child, reporting `table-transform-grouped`; identity transforms retain flat output. [ADR0041](decisions/0041-shared-table-borders.md) records shared Layout resolution and partial native fidelity. See [the table contract](../api/table-import.md).


### Chart data projection boundary

Parser resolves separate effective chart trees and inventories auxiliary relationships. Its private data reader interprets dense caches/literals; the workbook adapter reuses bounded nested OPC inspection, validates local worksheet data and reads stored values without evaluating formulas or following external links. [ADR 0002](decisions/0002-chart-data-projection.md) assigns complete-cache precedence, missing-cache recovery, explicit disagreements and per-import work/expanded-byte accounting. Every referenced cell read is charged to prevent amplification through shared ranges. Source part/identity/format provenance stays Parser-owned; Core stores semantic arrays/styles only.

Public structural reports clone MCE records before freezing, allowing later content-only chart roots to append private decisions safely. The final content report snapshots all decisions. Each frame maps independent Core data; fresh export rebuilds separate chart/workbook pairs with current category-aligned values, including explicit null gaps, and PresentationML transforms. Preview recomputes chart layout and is marked fallback. No-op retains the complete chart source; the unrelated text-run save profile also retains chart/workbook payloads and identities. Default editing of chart data/styles and session rebuild remain blocked. Separate explicit `chartEditSaveProfile: "existing-values-v1"` permits bound existing numeric values only: Parser stages lexical cache and XLSX worksheet changes, compares inner/outer allowed payloads, rebinds the actual candidate cache/workbook identities and values, then verifies full Core equality. Shared parts/cells, formulas/dependencies and unsupported normalized edits block. Nested validation shares outer session quotas; no source baseline advances. See [chart import](../api/chart-import.md), [numeric saves](../api/chart-value-save.md) and [ADR0034](decisions/0034-chart-numeric-preservation-save.md).

## Cached diagram preview

The separate [cached diagram preview](../api/diagram-preview.md) is opt-in and Parser-owned. DSP wrappers are translated with existing DrawingML readers into ordinary Core groups; provenance, unknown freshness and capability reports stay outside Core. No cache XML or SmartArt node model enters Layout/Renderer. Captured preview options propagate through Node imports and actual-candidate reprojection, while cached text remains unbound and preservation mutations stay blocked. Direct fresh export writes derived ordinary shapes, not native SmartArt. See [ADR0018](decisions/0018-cached-diagram-preview.md).

## Page Fade transitions

[ADR 0012](decisions/0012-slide-fade-transition.md) assigns immutable timing to Core, detached transport to Layout and native mapping to Exporter/Parser. Parser checks selected raw origins without globally broadening MCE namespace support. Source-only effects remain reported; current preservation plans reject transition mutations and retain unrelated source timing markup. Static SVG renders the final page; playback is separate.

The opt-in [direct diagram text observations](../api/diagram-text-comparison.md) remain Parser-owned. Exact paragraph disagreement is descriptive; it does not grant Core node ownership, native writes or cache-freshness certification. See [ADR0030](decisions/0030-diagram-text-observations.md).


Native timing source inspection is a separate opt-in Parser operation; [ADR0038](decisions/0038-native-timing-source-inventory.md) retains raw/effective identities without granting Core mapping, native writes or playback.


[ADR0039](decisions/0039-stable-paragraph-target-preservation.md) assigns opt-in raw paragraph-target proofs to Parser. The existing lexical source writer preserves target XML during scalar text edits; Core gets no format IDs or effects. Unknown/character targets retain guards and native producer observations do not grant playback admission.

[ADR0043](decisions/0043-source-table-style-projection.md) keeps source table-style GUIDs, relationship/default selection and conditional order in Parser. Target-scope RGB/font values become direct Core cells; effective provenance never grants raw source-mutation authority. Fresh materialization and source-preserving scalar saves remain distinct.

[ADR0049](decisions/0049-source-supplemental-font-projection.md) keeps raw font collections and source selection metadata in Parser until the native-slot cascade selects them. Only admitted per-script family strings enter Core; provenance and candidate reimport retain the captured profile. Existing scalar permissions do not allow font/map/style writes. Source-rule preservation and warned fresh materialization remain distinct.

[ADR0050](decisions/0050-notes-master-text-inheritance.md) keeps notes-master/type/theme identity in Parser behind an explicit notes text profile. Supported effective styles become ordinary Core notes; source saves retain the captured profile and complete snapshot equality. Native notes-page layout and source style writes remain outside this contract.

[ADR0051](decisions/0051-declared-slot-supplemental-projection.md) adds an explicit conservative v3 font profile. Only an effective native bucket/selected collection admits its supplemental rules; a Latin dictionary does not fill an absent EA/CS selection. Actual source entries remain reported unsupported with dependencies/targets, and candidates retain v3. No implicit Office font, host substitute, new Core field or source style write permission is inferred.


The Parser also offers an explicit [master picture source inventory](../api/master-picture-inventory.md). It reports effective raw-source picture identities, lexical geometry and relationships within listed masters, without mutating Core or granting edit/delete permission. Inactive MCE branches, unknown wrappers and complete reference closure remain separate work.

The optional `pictures-v1` content profile constructs verified listed master/layout/slide paths as shared Core definitions and projects ordinary picture fields and their containing group geometry through the existing bounded asset mapper. It reports omitted subtrees and uses the same asset identity across master, layout and slide owners. This is a partial fresh-output projection, not an editing binding or a source closure proof; it cannot be combined with current preservation edit profiles. See [content import](../api/content-import.md) for the exact boundary.

The separate `artwork-v1` profile adds a bounded plain-shape subset without changing `pictures-v1`. Source drawing order and shared ownership remain intact; omitted text, inherited paint, backgrounds and unknown properties remain explicit limitations.

`artwork-v2` further projects direct text in eligible shapes through the shared content reader. It does not create preservation bindings, resolve master text-style inheritance, or map fields and actions; source audits continue to report omitted nodes and properties.

`artwork-v3` resolves supported text theme references against the source master or layout owner, without fabricating a slide context. It materializes a single shared value and cannot represent page-specific overrides on inherited artwork; it still does not infer master text-style defaults or provide edit-save bindings.

`artwork-v4` applies the same owner scope and bounded supported color transformations to explicit shape fill and line paints. These become fixed Core colors; style matrix and per-page color changes are not inferred, and paint dependency provenance is not yet exposed.

`artwork-v5` also projects supported explicit solid slide/layout/master backgrounds into their respective owners. Source ordering and local background precedence are retained, while unsupported background declarations remain reported fallbacks and source-preserving editing remains separate.

`artwork-v6` uses the bounded DrawingML matrix interpreter to resolve eligible solid `bgRef` entries and placeholder color programs in the owning source scope. It does not retain a live matrix link or authorize source-preserving edits.

Exact custom decimal values use the Core canonical decimal-string variant and fresh `vt:decimal` output; the Parser offers separate read-only and source-value-edit opt-ins while preserving its existing default real-number behavior. An explicit structural-save opt-in also admits bounded decimal appends and custom-part creation; earlier profiles do not insert this new type. See [ADR 0066](decisions/0066-exact-decimal-custom-properties.md).
