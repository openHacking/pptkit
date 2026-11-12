# ADR0041: Shared physical table borders with sparse direct overrides

Status: accepted for direct styles, 2026-10-26. Full table-style inheritance, source writes and native fidelity remain open.

## Decision

Add sparse four-sided stroke overrides to Core cell styles. Normalize provided edges against the existing uniform stroke; preserve absence so existing tables keep their legacy path. Layout owns deterministic conflict resolution and returns a versioned physical edge graph plus exact cell geometry. It uses a bounded ownership grid and prefix geometry; request-wide budgets cover all opt-in tables. Core stays format independent and SVG/PPTX consume the same chosen styles.

Priority is visible, explicit, width, dash rank, then upper/left owner. Same-owner merge interiors are invisible. This deliberately small direct-style rule does not pretend to implement the Office theme/table-style cascade. A per-edge noFill cannot silently erase a neighbor's visible candidate.

## Native findings and representation

Independent native output exposed a missing-perimeter bug: writing only physical slot edges puts noFill on a merged anchor's right/bottom, which some producers treat as the logical merged border. Uniform logical outer sides now also go on the anchor. Continuations retain physical hints and Office-required duplicated merge spans. Heterogeneous sides get explicit native warnings and noFill anchors; neighboring/perimeter segments remain, without claiming native parity.

Office ignores graphic-frame rotation/flips. Fresh transformed tables use a native group holding the authored root identity/action and a private editable table child. This keeps the intended group transform in a supported OOXML location; it introduces a documented hierarchy change and does not certify all applications. Named LibreOffice PDF geometry/flip discrepancies remain partial.

Parser projects supported direct unequal/incomplete edges sparsely; full uniform merged sides can aggregate from continuation hints. Heterogeneous/missing segments remain source-only and original source bytes remain preserved. Legacy anchor-only merges keep their established projection. Source style and structural edits stay blocked.

## Alternatives, consequences and evidence

Independent per-cell CSS borders cause double/conflicting edges and disagree with fresh output. Collapsing unequal edges into one stroke loses authored data. Format rules in Core would violate package direction. The shared Layout graph avoids these alternatives without a new package or dependency. A complete theme/style cascade is future work.

See [API/evidence](../../api/table-borders.md). Core/Layout/Parser, exporter and cross-package tests verify direct rules, ownership, quotas and fresh/source-save boundaries. Independent Python/Chromium readers inspect physical OOXML rather than reusing the writer. PyMuPDF and PNG observations separate geometry/color coverage from strict width/alpha/transform fidelity, with an unedited source-control PDF. Native partial results and interactive/source-write gates are not promoted to pass.
