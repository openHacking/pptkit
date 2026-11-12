# ADR 0031: Named native cache conflict evidence and write admission

- Status: accepted research/policy boundary; public native SmartArt writes remain unadmitted
- Date: 2026-10-26

## Problem

A source data graph and its stored DSP cache can carry conflicting text without malformed XML or broken references. One native application's successful open/save cannot decide whether another application or PPTKit's static cache preview will display the same characters. Cache omission also does not supply PPTKit with generated preview geometry.

## Evidence

The [controlled protocol](../../testing/smartart-stale-cache.md) compares three exact layouts (`process3`, `orgChart1`, `cycle2`). LibreOfficeDev 26.8.0.0.alpha0 uses nonempty cached text in direct/reopened PDF and retains disagreement after saving; empty controls render original labels but retain empty caches. PowerPoint 16.104 build 16.104.25121423 displays source data text on open and writes rebuilt caches on save. All fifteen PowerPoint controls reopen. Three separate native layout-change/return cases preserve content identities/roles/order and source markers. Six additional source-only conflict controls extend data-only rebuilding to three rounds per family, retaining content identities and package/resource counts. Screenshots/AX and 24 independent ZIP/XML/native-output checks are retained; font/geometry/paint fidelity is not established.

The independently controlled mutation changes data text outside the native UI before opening. It is not native text-editing evidence and does not supersede the generated-output editing matrix in ADR0025/0026. The imported cache is unknown even when a native output's directly comparable character vectors agree.

## Decision

1. Preserve existing opt-in cached preview and report explicit direct paragraph disagreements through ADR0030; always retain unknown freshness for imported caches. Preserve default output and source-save permissions.
2. Reject any native SmartArt write admission based solely on PowerPoint regenerating stale or empty cache. LibreOffice and static preview results make that insufficient. A managed semantic edit needs accepted layout/text/style/resource generation plus a synchronized validated cache, or an explicit unavailable/blocked result appropriate to the operation.
3. Treat missing/empty cache as an explicit fallback state, not a cross-application fidelity strategy. Native regeneration does not produce geometry in the static importer, and observed producer serialization differs.
4. Keep data/layout/style/colors/cache/resources under a format-owned publication transaction. The private session protocol of ADR0029 remains a research component, not permission to change package parts.
5. Gate each layout and target application independently. Continue shared-source/image/rich-text/parameter/dependency variants, geometry/font/paint comparisons, actual native node-edit operations and complete source-save transaction admission before public family support.

## Validation and remaining work

The retained-output verifier uses public Parser inspections/no-op saves, then Python reconstructs content IDs, roles, paragraphs and ordered parent edges independently. It checks source/input hashes, actual input conflicts, cache output vectors and repeated-round resource counts. Negative controls reject byte snapshot mismatch, cache character corruption, and internally consistent content-ID rewriting. No native binary or third-party source is redistributed by this decision. ADV-03.7 and its family dependencies remain open; the completed matrix narrows the remaining evidence rather than closing those parents.
