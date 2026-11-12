# Multi-master native admission and ID collision fix — 2026-10-27

PowerPoint for Mac 16.104 exposed a real failure after the automated [authoring foundation](multi-master-authoring.md): its original public-authoring source triggered a repair prompt. Repair was cancelled; the source, warning AX text and screenshot remain under `.tmp/multi-master/native-b85qrfu0`. Earlier structural and browser passes did not establish native admission.

## Differential evidence and production correction

The [independent manifest](multi-master-native-evidence.json) hashes all five input/open observations and records this two-factor control:

| Source | Unused master | Master IDs overlap layout IDs | First native open |
| --- | --- | --- | --- |
| Original public output | Present | Yes | Repair prompt; cancelled |
| Public output omitting unused declaration | Absent | Yes | Repair prompt; cancelled |
| ID-only patched original | Present | No | Normal four-slide window |
| ID-only patched no-unused source | Absent | No | Normal four-slide window |
| New public output after production correction | Present | No | Normal four-slide window |

Each patched pair changes only the presentation part's master ID values; every other ZIP payload is exact. The actual production correction allocates additional master IDs **after the entire layout ID range**. The legacy default master/layout IDs and old-input path remain unchanged. It is a deterministic allocation policy derived from the named native behavior, not a claim that a general XML/schema validator proves all native rules.

The new unpatched public source differs from the original source in `ppt/presentation.xml` master IDs only. Its SHA-256 is `943270e78b2fe373a20e2983c239b22211ee4c973bf9fa9d23f5b64fa0576e89`. Both the previous failed prepared run and all differential inputs remain preserved. [Exporter regression tests](../../packages/pptx-exporter/test/masters.test.mjs) check the **union** of master and layout IDs, including 24 interleaved/unused layouts; the previous check tested only layout uniqueness and missed the cross-kind collision. The independent browser DOM check also verifies the union.

## Actual save, close, reopen and output

Starting from `fixed.pptx`, PowerPoint performed three separate Save As operations to round1/round2/round3, closed each document, and opened the exact saved path before capturing a four-page slide-print PDF. AX and screenshot files for each reopened path remain alongside the PPTX/PDF files. The first-open PDF has four pages; the three reopened PDFs have twelve pages. No repaired document is counted as a positive source.

[The Python oracle](../../scripts/research/oracles/multi-master-native.py) uses stdlib ZIP/XML and independent PyMuPDF/Pillow checks, with no production Parser/Core dependency. It follows presentation order and actual slide→layout→master→theme and notes backedges, then maps content by stable native identity. PowerPoint renames/reorders some layout/theme parts and relationship IDs; direct filename comparison would incorrectly report cross-domain theme changes.

All four registered masters (including unused), five layouts, native IDs, full theme trees, native master/layout ownership, declared and inherited backgrounds, slide text/geometry/styles and notes content pass canonical comparison. Content types and complete relationship edges are checked too. The narrowly recorded native transformations are: paragraph `indent=-0` to `0`; ordinary `noGrp=1` locks; slide-image `noGrp/noRot/noChangeAspect=1` locks; and the exact empty slide-guide extension. Other unknown nodes/attributes remain compared. This is an explicit comparison profile, not literal source XML equality or unqualified editing equivalence.

Presentation/view properties and app/core metadata changes are stored as full before/after XML in the manifest. After the first native save, every payload except `docProps/core.xml` is byte-equal across rounds 2 and 3. Raw native files remain intact.

All twelve reopened pages match the first native PDF in full 144dpi raster bytes, glyph trace and embedded font-program hashes. Expected actual faces are Georgia, TimesNewRomanPSMT, Georgia and ArialMT. Red/blue/red/default text and inherited background colors, text order and visible ink are checked. Native PDF color floats render the green channel of `CC2200` and `112233` one byte lower in MuPDF; ink detection uses the actual PDF float conversion rather than silently demanding different source RGB. Declared text colors and exact native-to-native raster comparisons remain strict. This does not establish browser/native pixel equality or general font equivalence.

The generated four-page native contact sheet was visually inspected: correct red/blue/red/default backgrounds and labels are visible without clipping. The contact image and each PDF have hashes in the manifest. The PDFs are verification captures rather than authored deliverables.

## Actual failure controls and current consumers

Ten actual corrupted PPTX files reject: cross-kind ID collision, wrong layout owner, wrong master theme, color, font, background, slide text, notes, removed unused registration and content type. Two actual PDF corruptions reject: a white overlay covering the first label and swapped page order. Files and exact reasons are retained; the oracle does not only alter an expected JSON report.

[Current captured-input runner](../../scripts/inspect-native-multi-master.mjs) checks all four source/native SHA values, full structural inspections and reports, and exact no-op bytes in Node, browser ESM and global with zero network requests. Its [manifest](multi-master-native-parser-evidence.json) contains twelve consumer no-ops and full snapshot hashes. It starts no native app and does not grant a source master/style writer or authored registry import.

## Remaining scope

The failure correction and no-edit native admission are complete for this fixture/application. [Named actual layout/text/selected-theme controls](multi-master-operations.md) now provide nine further reopen rounds. Complete master artwork, shared-definition mutation, generalized template operations and multi-master/notes/field parent gates remain open. Neither a preserved unused master nor a stable no-edit PDF proves those operations. No source master/style write permission was added.


## Final automated validation

`pnpm build`, `pnpm typecheck`, complete `pnpm test`, lint/boundaries/docs examples and diff checks pass. The complete test command includes Core150, Layout139, Parser696, Exporter57, Renderer19, CLI11, 34 visual tests, 118 import integration tests and isolated installed runtime/types consumers. Logs are `.tmp/masters-native-{build,typecheck,full-test,lint,docs}.log`. Production code was fixed before the complete run. The independent native oracle and captured-input runner completed against their final manifests; all referenced source/native hashes and snapshot hashes were rechecked. Existing dev-app large-chunk build advice remains, with no performance admission claimed.
