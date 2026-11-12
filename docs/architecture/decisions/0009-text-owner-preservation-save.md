# ADR 0009 — Explicit shape-text and notes ownership in preservation saves

Status: accepted for experimental `text-runs-v2`.

Subsequent [ADR 0010](0010-table-text-preservation-save.md) adds an explicit V3 table-anchor character profile without changing the V2 scope/report shape.

Date: 2026-10-25.

## Decision

Extend the [existing preservation pipeline](0008-direct-text-preservation-save.md) through a new explicit profile, keeping V1 and default runtime report shapes compatible. V2 permits fixed-topology scalar replacements in bound shape text and direct notes bodies as well as text boxes. It shares the existing character primitive, ZIP writer, candidate comparison/reprojection, budgets and lifecycle. No new dependency, package, Core/IR field or public low-level XML writer is introduced.

Drawing provenance keeps required Core element identity; V2 adds an optional owner tag. Notes have a separate public record with slide identity and source body/part identity, never a fabricated Core element. Private binding analysis unifies scalar paths while keeping notes and drawing identities distinct. The whole normalized document remains the unit of mutation admission and candidate verification.

Notes binding requires exactly one direct body placeholder and a local text body. Master/inherited prompts, ambiguous bodies, table/chart/SmartArt text and unmapped shapes acquire no permission. Duplicate notes-part owners keep the structural import rejection. Raw guards scan selected slide and notes parts, including inactive branches, keyed by part plus numeric source object identity. Unrelated equal IDs do not cross-block; edited source owners with range/opaque targets block atomically.

Derived plainText is exempt only at actual normalized text owners. Custom data receives ordinary diff analysis. Source bindings across drawings and notes share one report-count budget; complete analysis, output comparison and actual-candidate reprojection retain the existing limits.

## Evidence and consequences

The [public owner contract](../../api/text-owner-save.md) links dialect/owner/negative tests, three-round mixed-deck tests, Node/browser byte parity, independent DOM indices, installed consumers and native headless evidence. The native producer retains edited shape/notes text but drops unknown content retained by PPTKit; this is producer loss, not a fidelity pass.

V2 advances IMP-09.2.1.2 while its inherited/native gates remain open. It does not add range remapping, arbitrary notes-page editing, property/topology changes, rebasing, native interactive/PowerPoint guarantees or general rebuild.
