# ADR0039: Stable paragraph-index targets during scalar source edits

Status: accepted for explicit structural source preservation, 2026-10-26. Native playback admission remains deferred.

## Problem and decision

Existing fixed-topology text profiles guard all subtargets. A character range can become wrong when text length changes. A paragraph range can retain its source identity when the direct paragraph sequence and all surrounding XML remain unchanged. Admit this distinction through `textTargetSaveProfile: "paragraph-stable-v1"`, requiring an existing text-run profile; default behavior stays compatible.

The primary [ParagraphIndexRange definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.paragraphindexrange?view=openxml-3.0.1) defines start/end paragraph numbers and illustrates the first three as `0..2`. The [ShapeTarget definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.shapetarget?view=openxml-3.0.1) places the text target on a shape. We infer only structural index preservation under an unchanged paragraph sequence, not executable/native semantics.

Parser proves raw owner uniqueness across all branches, direct body/paragraph correspondence and source text atom ancestry. Only transparent, in-bounds `spTgt/txEl/pRg` forms bypass their individual range guard. Other references remain independently blocking. Notes/table/diagram/chart/group-level targets and character/unknown ranges do not gain authority. No timing bytes, Core animation contracts or package dependencies change.

Reuse the existing source lexical writer, whole-candidate diff, reservations, package comparison and actual Core reprojection. Proofs are private transient maps, bounded across all parts with cooperative cancellation. Numeric identity lookup follows the existing content profile rather than silently adopting the separate inventory's canonical-only policy.

## Alternatives and risks

Keeping every paragraph target blocked would unnecessarily prevent a structurally provable edit. Automatically broadening all existing profiles would change documented authority. Generic range rewriting would require unresolved Unicode/native semantics and source structural policies. A second document model or raw AST public API would add ownership and transport obligations without helping this fixed-topology transformation.

Do not infer paragraph moves from strings or treat an unchanged range as playback proof. Raw inactive aliases can make an effective owner ambiguous; they must defeat the proof. A supported target must never cancel a separate guarded target. Paragraph insert/delete/reorder/topology and native effects remain separate tasks.

## Evidence and admission boundary

See the [contract/tests](../../api/paragraph-target-save.md) and [independent/native manifest](../../testing/paragraph-target-save-evidence.json). Three-round source edits pass exact Python ZIP/XML checks and independent browser DOM checks. The named LibreOffice resave retains text but rewrites the sample range `0..2` to `0..0`, also on the unedited control; native range preservation is partial. Saved-output PDF reopening is not playback. ADV-07.8 and ADV-06 native gates remain open.
