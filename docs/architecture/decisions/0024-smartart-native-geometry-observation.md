# ADR 0024 — Observe native diagram geometry before implementing layout

Status: research, advancing ADV-02.1.3.1; no family/layout admission.

## Problem and decision

The six standard generated diagrams from [ADR0023](0023-native-smartart-generation-research.md)
establish scoped producer text observations. They do not define node/label/arrow
geometry, fitting behavior or results for empty/single/long cases. A deterministic
PPTKit layout needs independently measured expectations and explicit failure
cases before its geometry can be evaluated.

Use native flat-ODF and PDF artifacts as **separate observation surfaces**.
The [optional runner](../../../scripts/inspect-smartart-native-geometry.mjs)
renders every presentation-mode candidate directly to FODP and PDF. It validates
the full three-family × four-case × two-cache manifest, source hashes and
unchanged inputs, then records geometry and text without declaring compatibility.
This is an experiment over local fixtures, not a public ODF/PPTX importer.

The [independent Python extractor](../../../scripts/research/extract-native-diagram-geometry.py)
locates a diagram group using the source graphic-frame's explicit name. Missing
and ambiguous groups stay distinct from an empty drawing. Only the drawing page
is inspected; masters and notes cannot supply a false group. Occurrence ordinals
retain repeated names. Stored paragraphs include spans, explicit spaces/tabs/soft
breaks, frame text boxes and nested lists. This is stored plain-text observation,
not a full ODF text/list semantic projection.

The [pure observation module](../../../scripts/research/diagram-geometry-observation.mjs)
converts explicit untransformed x/y/width/height lengths to page points. Python
Decimal expectations independently check every resolved box within 1e-8 points.
Frame containment uses an explicit 0.1-point tolerance; this is a bounding-box
finding, not proof that glyphs fit. Units are pt/pc/in/cm/mm and unitless zero.
Missing, unknown and negative-size geometry remains incomplete. Shape/group
transforms, mirrors and enhanced text-area expressions remain raw. Implementing
their semantics would require a separate bounded adapter and evidence under the
[ODF drawing schema](https://docs.oasis-open.org/office/OpenDocument/v1.3/os/part3-schema/OpenDocument-v1.3-os-part3-schema.html);
this observation deliberately does not infer them.

## Identity, comparison and resource boundaries

Producer drawing names may match presentation model IDs, but they are **evidence,
not stable semantic identities**. Cycle's native output repeats an arrow name.
The report retains all source-index/name evidence and occurrence ordinals, and
reports missing/unmatched/non-presentation/ambiguous source identities separately.
It never associates a drawing to a data node by visible text or array position.

Controlled absent-cache/empty-cache comparisons use drawing order, explicit source
associations, resolved point boxes and raw transformed frames/transform/paragraph/
text-area/mirror records. Random unmatched producer names and source XML ordinals
are excluded. Equality is exact observed equality under that policy; it does not
prove cache freshness, cross-platform determinism or semantic/visual equivalence.

The pure module bounds shapes, cumulative work and captured text. Repeated names
and ambiguous matches charge expansion work, preventing quadratic report growth
from escaping the budget. Invalid input/limits, source-index collisions and
exhausted budgets throw without returning a partial success. Results are detached
and ordered by source occurrence. The optional extractor separately limits UTF-8
XML bytes/nodes, drawing count, paragraph work/text and explicit whitespace counts;
DTD/entity declarations reject. It is an offline research oracle, not a hardened
general-purpose format parser.

## Evidence and remaining work

The [24-case baseline](../../testing/smartart-native-geometry-evidence.json)
records raw source hashes, native/PDF/picture hashes, group association, raw and
resolved drawing rectangles, font-face declarations, PDF span/font/paint bounds,
stored drawing text and independently extracted PDF text. Paths/spans are
separate observations: no invented shape ownership or clipping interpretation is
attached to them. The native renderer is the named LibreOfficeDev development
runtime; no PowerPoint interaction is established.

Long-text cases expose a real remaining gate. Cycle's inspected PDF picture clips
the long top label; PDF text also has CJK/extraction discrepancies. Stored drawing
text and PDF-extracted text are compared separately, so rendering failures cannot
be misreported as source graph loss. Standard geometry or mere frame containment
must not admit long-text layout. Empty graphs have no authored markers: lexical
success for zero markers is vacuous and the report retains markerCount.

All 24 native drawing snapshots retain the authored text under word-spacing
matching; six long-case PDFs have unmatched markers. Twelve controlled absent/
empty-cache pairs have exact observed geometry equality under the stated policy.
There are 124 recognized drawing occurrences: 112 explicit axis-aligned rectangles
and 12 transformed arrows retained as source-only. Process standard/long cases
each have three boxes outside the source frame under both cache policies. The
[three picture reviews](../../testing/smartart-native-geometry-review.json)
keep manual visual findings separate from automated results: two Cycle images
show the standard/long distinction, and one Process image confirms the right card
extending beyond the frame. Matching cache controls do not make these failures pass.

Seven portable Node groups cover unit conversion, frame tolerance, detached
output, unsupported/ancestor transforms, source-name ambiguity/repetition, budgets,
invalid inputs and cache comparison mutations. Seven optional Python fixture
tests cover XML extraction, text boxes/lists/whitespace, ancestor transforms,
master/group scoping, DTD/resource rejection and Decimal expectations. Live
producer observations are not replayed by default CI. See the
[reproduction protocol](../../testing/smartart-native-geometry.md).

Keep Core/Layout/Parser/Exporter/Renderer public contracts unchanged. Next work
remains width/depth/assistant variations, different frame sizes and direction/
theme parameters, actual text-area/glyph fitting and font strategy, deterministic
layout per exact family, valid/stale cache experiments and PowerPoint node edits.
Neither ADV-02.1.3.1 nor any parent SmartArt task is completed by this baseline.


Subsequent attribution controls and generated-output PowerPoint edits are recorded in [ADR0025](0025-generated-smartart-native-edit-evidence.md). Original-source Process also overflows in the named LibreOfficeDev controls; generator-only attribution is unproven. This does not erase the recorded geometry failure.


[ADR0028](0028-cycle-layout-feasibility.md) adds an owned deterministic Cycle geometry experiment and explicit PowerPoint circle-cache comparisons. Its 0.1-point research threshold is not met by the six retained cases; original LibreOffice observations and failures remain separate.
