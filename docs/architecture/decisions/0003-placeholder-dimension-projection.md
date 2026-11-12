# ADR 0003: Source placeholder ancestry and dimension projection

Status: accepted for the bounded experimental Parser profile; full template semantics remain pending.

Date: 2026-10-25.

## Problem

The content reader previously required complete local geometry. Source placeholders often carry only content or a subset of their geometry, with dimensions supplied by layout/master placeholders. Mapping names, copying master text, or putting raw template XML into Core would create false relationships and duplicate content.

## Decision

Parser owns source placeholder ancestry and dimension provenance. It matches slide/layout placeholders by numeric index and layout/master placeholders by an explicit supported type mapping. Missing and ambiguous matches do not pick a fallback candidate. Multiple source masters stay independently scoped by verified relationships.

The resolver reads the retained effective MCE trees. Each part is scanned and each parent graph resolved once. Lookups use index/type maps. Ambiguity reports contain a candidate count and matching key, avoiding a candidate-list expansion for every child. A distinct-declaration budget and cooperative checkpoints complement existing ZIP/XML/OPC/report limits.

Resolved offsets/extents retain per-property EMU values and source locations in `ImportReport.placeholderInheritance`. The content projection materializes local Core boxes and supported preset names. It does not create Core templates or copy master/layout prompt text. A nearest explicit unsupported/custom geometry choice blocks an older preset from taking effect. A separate bounded frame/paragraph/list/run text profile is documented in [ADR 0004](0004-text-style-projection.md); it does not imply theme, paint, or complete Office style inheritance.

Core/Canonical IR stay unchanged. This is a source import projection, not a replacement for the Core multi-master/override/rebinding contract required by AUT-07. That contract must be designed explicitly before IMP-05.1 is closed. Source editing remains governed by IMP-07/09, not by a successful fresh export.

## Implementation evidence and unsettled semantics

The matching and per-dimension policy was checked against the primary [python-pptx placeholder implementation](https://github.com/scanny/python-pptx/blob/c38d5f5c6850ae3aefdc3a86dbf9bd0af35cf346/src/pptx/shapes/placeholder.py), in particular `_BaseSlidePlaceholder._base_placeholder`, `LayoutPlaceholder._base_placeholder` and `_InheritsDimensions`. This reference provides implementation evidence, not an Office interoperability certification.

Transform behavior is not sufficiently established by those dimensions. The primary [Apache POI implementation](https://github.com/apache/poi/blob/a70c44e27644871afaf44bc696e0c3c714d65cf0/poi-ooxml/src/main/java/org/apache/poi/xslf/usermodel/XSLFSimpleShape.java) fetches a whole ancestor `xfrm`, whereas python-pptx's inherited-dimension implementation is property-based and its basic rotation access follows local transform state. These approaches do not independently prove identical partial-transform results. Consequently the profile reports inherited non-neutral rotation/reflection as unsupported until explicit native fixtures establish the contract; local complete transform behavior remains supported.

Grouped placeholder coordinate mapping and header/slide-image master matching also remain unsupported. Unsupported declarations and candidate counts stay visible in provenance; they are not silently removed from source no-op saves.

## Verification and remaining gates

See [the public contract and tests](../../api/placeholder-inheritance.md). Tests cover independent source XML, defaults/lexical aliases, explicit zero/false, duplicate indices/types, multiple masters, opaque extension islands, MCE/aliases/dialects, malformed inputs, cancellation/budgets, three fresh rounds, browser and installed public consumers. Text cascade evidence is listed separately in [its contract](../../api/text-style-inheritance.md).

Native layout/template editing, complete property/theme inheritance, dynamic rebinding and preservation writes remain separate gates. No IMP-05/AUT-07 parent or child is closed by this bounded result.
