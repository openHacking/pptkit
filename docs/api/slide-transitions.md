# Experimental slide transitions

`slide.setTransition()` replaces or clears an immutable page transition. `addSlide({ transition })` accepts the same input; slide copies retain independent snapshots. This is implementation progress on ADV-05.1.1. Native interactive settings/playback and preview playback remain open. Push/Wipe data semantics are described in the extension below.

<!-- doc-test: docs/examples/slide-transitions.ts -->
```ts
import { createPresentation } from "@pptkit/core";
import { generatePptx } from "@pptkit/pptx-exporter";

export async function createFadeDeck() {
  const deck = createPresentation();
  deck.addSlide().addElement({
    type: "text", content: "First page",
    box: { x: 40, y: 40, width: 400, height: 60 },
  });
  const second = deck.addSlide({
    transition: { type: "fade", durationMs: 1234, onClick: true, afterMs: 5000 },
  });
  second.addElement({
    type: "text", content: "Fade into this page",
    box: { x: 40, y: 40, width: 400, height: 60 },
  });
  return generatePptx(deck);
}
```

## Model and timing

`SlideTransitionInput` admits Fade, Push and Wipe, with required `durationMs`, optional `onClick` and optional `afterMs`. Push/Wipe additionally require cardinal `direction`; see the extension below. Duration describes the visual transition **into** the owning page; click and automatic timing govern advancement **from** that page. These independent concepts share the native slide transition record. Durations/delays are integer milliseconds in `[0, 4294967295]`. Zero is explicit and retained. Omitted `onClick` normalizes to `true`; omitted `afterMs` disables automatic advance. `onClick: false` without a timer is allowed for navigation controlled by other actions. When click and automatic timing coexist, the consumer/native slideshow controls which event advances first.

Core validates unknown fields/effects, invalid booleans, fractional/nonfinite/negative/overflow times with slide/path diagnostics. Setters capture immutable input before replacement; a failed capture leaves the previous snapshot intact. Core and Layout own no timers, playback events or navigation. A transition-free slide retains the earlier IR shape. The optional additive IR v2/Layout field is detached, serializable and preserved through duplication and table continuation metadata. Custom `PresentationSlide` implementations must implement the new getter/setter.

## Native output and reading

Fresh output emits `p:transition/p:fade`, `advClick` and optional `advTm`. Exact milliseconds use `p14:dur` in the Office 2010 namespace with a local `mc:Ignorable="p14"`; readers without that extension use the coarse `spd` fallback (`fast` through 750 ms, `slow` from 1500 ms, otherwise `med`). These fallback thresholds are PPTKit policy, not exact duration equivalence. See Microsoft's [duration attribute](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.transition.duration) and [transition extension/timing example](https://learn.microsoft.com/en-us/office/open-xml/presentation/how-to-add-transitions-between-slides-in-a-presentation). No transition data is duplicated into PPTKit custom metadata.

`inspectPptx` exposes optional per-slide `PptxTransitionInfo`: mapped/source-only status, speed, onClick, optional afterMs/durationMs/type/reason. Both structural and content imports map a selected plain Fade/Push/Wipe with explicit `p14:dur` and explicit direction for side effects. The reader resolves the selected node's verified raw origin to inspect this attribute without claiming to understand every `p14` MCE branch. It never borrows values from inactive Choice branches. A selected speed-only fallback remains source-only because exact milliseconds cannot be inferred safely.

Through-black fade, effects beyond Fade/Push/Wipe, sounds, unknown attributes/subtrees and extensions remain source-only with capability findings. Malformed supported scalar values and duplicate transition/fade declarations reject with a source part. Unknown source content survives unchanged no-op and unrelated supported character edits; changing the Core transition blocks all current edit-save profiles. Fresh export of a partial source-only projection omits those unsupported transitions; use the preservation session to retain them.

## Static preview and evidence

SVG renders the complete final page, independently of transition timing, with identical SVG bytes before/after transition removal. There is no implicit animation, hidden first frame, CSS timer or `reduced-motion` API in this slice. Playback is a separate ADV-08 task.

[Core tests](../../packages/core/test/transitions.test.mjs) cover snapshots, copy/clear/defaults, capture failure, located validation and numeric boundaries. [Parser tests](../../packages/pptx-parser/test/transitions.test.mjs) cover both dialects, exact no-op bytes, stale plans, unknown/raw/MCE counterexamples and malformed values. [Integration](../../scripts/transition-parity.test.mjs) checks three fresh rounds, detached Layout state, exact retained transition markup across unrelated text saves, static SVG invariance and independent Chromium DOM/native-attribute evidence. Packed Node/type consumers exercise public contracts. These automated tests do not certify native playback, interactive editability, repair-dialog absence or universal duration interpretation. ADV-05.1.1 and inherited gates remain unchecked.


## Native headless observations

The optional [runner](../../scripts/inspect-slide-transitions.mjs) and [evidence](../testing/slide-transition-evidence.json) record LibreOfficeDev 26.8.0.0.alpha0 on macOS saving three named samples and reopening its saved PPTX to PDF. Independent Python ZIP/ElementTree inventories both source and native transition branches. The producer retains 1234/2000 ms duration, changes zero duration to 10 ms, truncates 3456 ms automatic advance to 3000 ms, and drops both `advClick="0"` restrictions. The native output creates Office2010 Choice and speed-only Fallback records. Parser selects Fallback and reports source-only without inventing exact duration. This is observed data drift, not a native timing-fidelity pass or actual playback evidence.

Reproduce after building: `node scripts/inspect-slide-transitions.mjs [soffice-executable]`. Each run creates an isolated profile and directory and records source/native/PDF hashes and command output. Successful PDF reopening establishes file processing only; interactive settings, repair dialogs, actual slideshow timing and PowerPoint remain untested.


## Push and Wipe extension

The experimental profile also admits `{ type: "push" | "wipe", direction: "left" | "right" | "up" | "down", durationMs, onClick?, afterMs? }`. Direction is the movement/reveal direction, not the starting edge. Native output maps these to `l/r/u/d` on `p:push` or `p:wipe`; see Microsoft's [Push contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.pushtransition) and [Wipe contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.wipetransition). Each of eight variants is independently tested. Fade rejects a direction field; side effects require one. Parser requires explicit native direction and exact duration, leaving a missing direction source-only rather than guessing. Unsupported directions reject with source location; ignored raw side attributes/children prevent mapping.

The existing lifecycle, static SVG, precision fallback and preservation restrictions apply to all three effects. Three fresh rounds and independent Chromium DOM checks verify every native direction. Installed Node/declaration consumers cover public exports. Playback/interruption and native application acceptance for Push/Wipe remain open under ADV-05.1.2; the earlier native evidence covers Fade only.


### Native observations for all eight side variants

Run `node scripts/inspect-slide-transitions.mjs --all [soffice-executable]` for the expanded 11-page experiment (three Fade plus eight side variants). [Side-transition evidence](../testing/side-transition-evidence.json) records independent Python source/native attribute inventories and saved-output PDF reopening with all 11 control texts. LibreOfficeDev retains 765 ms duration for all side variants, but changes 2345 ms advance to 2000 ms and enables click advancement. Push left/right native letters swap (`l→r`, `r→l`); Push up/down and all Wipe letters are unchanged. This records XML differences only, not observed slideshow movement. The earlier Fade-only evidence remains separately retained.

All selected native speed-only Fallback records remain source-only in Parser. No inactive Choice duration is borrowed. Neither directional fidelity nor timing/click fidelity is established by a successful native conversion. Actual native playback/settings, interruption, PowerPoint and dynamic preview remain open.
