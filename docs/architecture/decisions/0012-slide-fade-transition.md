# ADR 0012 — Page transition semantics without playback state

Status: accepted for experimental Fade authoring/import/fresh output. Native playback gates remain open.

Core owns one optional format-independent `SlideTransitionInput` snapshot per slide, containing the named effect and integer milliseconds/advance settings. The managed setter replaces or clears the snapshot atomically. Normalization materializes click defaults only when a transition exists; prior documents/IR remain unchanged. Duplication snapshots it independently. Layout transports detached semantics and does not evaluate clocks. Static SVG displays the complete page, independent of transition data. Preview playback/navigation will consume detached data through a separately tested adapter.

Exporter serializes native slide-level Fade with `p14:dur` for exact milliseconds and `spd` as a documented coarse compatibility fallback. Advance timing and visual duration stay distinct. It does not put transitions into custom slide metadata. Parser maps only a selected plain Fade with exact duration, checks the selected node's raw origin and never infers a duration from speed or an inactive MCE branch. It does not globally declare the Office extension namespace understood. Unknown effects, sounds, through-black and extensions stay source-only; malformed mapped scalar fields reject.

Current preservation profiles allow unrelated character edits while retaining original transition markup. New/changed transition values are whole-plan unsupported edits and cannot be approved through loss consent. This avoids adding a source attribute writer before its minimum-difference/reference rules exist. Both structural/content projection include located per-operation findings.

[The API and evidence](../../api/slide-transitions.md) document numeric boundaries, static behavior, independent DOM verification, three fresh rounds and retained unknown transitions. Native interactive settings and playback remain untested; no effect beyond Fade or native support certification is admitted. This advances ADV-05.1.1 without closing its application or inherited acceptance gates.


## Push and Wipe extension

The experimental profile also admits `{ type: "push" | "wipe", direction: "left" | "right" | "up" | "down", durationMs, onClick?, afterMs? }`. Direction is the movement/reveal direction, not the starting edge. Native output maps these to `l/r/u/d` on `p:push` or `p:wipe`; see Microsoft's [Push contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.pushtransition) and [Wipe contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.wipetransition). Each of eight variants is independently tested. Fade rejects a direction field; side effects require one. Parser requires explicit native direction and exact duration, leaving a missing direction source-only rather than guessing. Unsupported directions reject with source location; ignored raw side attributes/children prevent mapping.

The existing lifecycle, static SVG, precision fallback and preservation restrictions apply to all three effects. Three fresh rounds and independent Chromium DOM checks verify every native direction. Installed Node/declaration consumers cover public exports. Playback/interruption and native application acceptance for Push/Wipe remain open under ADV-05.1.2; the earlier native evidence covers Fade only.
