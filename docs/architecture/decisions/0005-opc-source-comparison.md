# ADR 0005 — Independent OPC source comparison profiles

Status: accepted for the implemented bounded profiles; broader semantic comparison remains planned.

Date: 2026-10-25.

## Problem

Core projections intentionally omit unsupported source content. Comparing two projected Core documents can therefore agree after unknown XML, unselected compatibility branches, binary resources or relationship semantics have disappeared. Comparing writer-produced values through the same projection is insufficient evidence for preservation writes.

## Decision

Parser owns `compareOpcPackages` above bounded OPC inspection and below presentation-specific semantic comparison. It does not import Core documents or depend on Exporter. No new package or dependency is introduced.

The default `opc-payload-v1` profile compares exact logical part names, content-type spellings, all expanded source payloads and source-local relationship identities. Only ZIP container representation is excluded. Exact byte comparison avoids relying on ZIP CRC32 as proof of binary equality.

The explicit `opc-xml-v1` profile permits a defined list of XML representation differences while retaining prefixes, namespace bindings, text whitespace, comments, processing instructions, child order and all MCE branches. It is deliberately more conservative than a schema-aware presentation comparator. Unknown QName-valued content prevents globally discarding namespace declarations.

The private XML reader retains a full document alongside its existing root when retention is requested, including prolog/epilog and declaration standalone state. OPC retains those documents in private maps, including the content-types stream. Existing public XML/OPC inspection types and Core/IR are unchanged. Iterative token traversal avoids recursive stack limits.

Both inputs/options are captured before yielding. Per-input inspection limits and a combined expanded allowance bound retained input; comparison work and retained findings have independent limits. Findings are frozen and deterministic. Truncation cannot imply equality or bypass candidate validation. Parsing failures and cancellation return no partial result.

## Consequences and admission gates

This provides an independent source comparison layer for no-op verification and later save validation. It can reject changes to source content even when effective visual values agree, which is intentional for preservation assertions.

It does not establish application schema validity, cryptographic signatures, visual equivalence, native editing/playback, or an edit-save reconciliation implementation. Nested ZIP/workbook parts are exact binary payloads rather than recursively interpreted structures.

Fresh-output effective-value comparison and source identity transformations need separate named rules and independent expected values. Renames cannot become acceptable until a full identity bijection and reference proof exist. IMP-08.1.1–.1.3 remain open for that wider scope.

Evidence: [API profile and test links](../../api/package-comparison.md). Browser DOM and independent ZIP/SHA256 oracles supplement Parser tests; a correct writer is not the only source of expected results.
