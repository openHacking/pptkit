# ADR 0023 — Minimum native SmartArt generation remains research

Status: experimental research; no public authoring or native-edit admission.

## Problem and ownership

Source-file editability does not establish that PPTKit can create a native diagram.
The next experiment must generate fresh content/transition/presentation identities
inside a fresh PPTKit deck, then independently inspect its structure and observe
the consuming application's output. Parser inspection, native display, layout
fidelity and interactive node editing are separate gates.

The pure [specimen writer](../../../scripts/research/native-diagram-fixture.mjs)
validates an ordered forest and serializes exactly process3/orgChart1/cycle2.
Its [presentation mapping](../../../scripts/research/native-diagram-presentation.mjs)
is a bounded, exact-layout experiment, not a DiagramML layout interpreter.
The [package builder](../../../scripts/research/native-diagram-package.mjs) owns
only a fresh package with a control text object. No package imports these scripts;
Core receives neither OOXML identities nor borrowed layout payloads. Public
authoring identities, rich text, geometry and preservation writes remain open.

## Candidate contract

- Explicit content/assistant nodes, roots, sibling order, plain paragraphs and
  Cycle's layout-generated closure. Cycles, multiple parents, unresolved children,
  assistant roots/descendants, nested Cycle input and unknown fields reject.
- Independent globally unique int32 point/connection IDs, reciprocal transition
  bindings and optional exact-layout presentation points/edges. Source and
  presentation associations can legitimately differ for Process detail text.
- UTF-8/XML-safe text with preserved CR, explicit soft breaks and rejection of
  illegal scalars. Work, depth, nodes, text bytes and total emitted XML are bounded.
- Explicit absent/empty cache controls only; no invented or stale shape geometry.
  Four external component inputs are offline snapshots, root/type checked and
  source-hash pinned by the runner. Resource-dependent payloads are outside this
  profile. These checks do not establish complete schema validity.
- Borrowed layout/style/color components are local research inputs. Licensing,
  redistribution and a production component strategy must be decided separately.

## Observations and correction

An initial word-spacing check reported every authored marker unmatched. It was
incorrect to describe this as absent node text. Controlled direct-PDF and
save/reopen-PDF observations reveal words fragmented across extracted lines.
The [lexical evidence helper](../../../scripts/research/pdf-text-evidence.mjs)
now keeps strict word-spacing and whitespace-insensitive matches separate. Neither
proves distinct occurrences, legibility, geometry or editability.

The first writer also set `custT="1"` on plain authored text. Microsoft's
[PropertySet documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.diagrams.propertyset?view=openxml-3.0.1)
identifies this as a text-customization flag affecting automatic formatting.
Plain paragraphs carry no custom formatting in this experiment, so the writer
now omits it and leaves native automatic fitting available. No placeholder text
or fixed font size was added to conceal the problem.

The [Cycle differential](../../testing/smartart-native-differential-evidence.json)
tests five variants, directly and after native save/reopen: unlocked text,
custom-text flag only, equivalent GUID IDs, source shell/generated data, and
fresh shell/source data. Only the custom-text variant requires whitespace removal
to match its markers; the other four match with word spacing. This supports the
flag correction in the named runtime, without proving identical geometry or
generalizing to every family. PNG artifacts and hashes permit picture review.

The [revised presentation candidates](../../testing/smartart-native-presentation-candidate-evidence.json)
cover three layouts × four cases (empty/single/standard/long) × two cache policies.
Independent Python ZIP/XML checks and exact no-op checks cover all 24 packages.
[Native producer observations](../../testing/smartart-native-presentation-producer-evidence.json)
cover the six standard cases: all save/reopen with all word-spacing markers and
the data inventory retained. Cache shape counts remain zero. LibreOfficeDev
26.8.0.0.alpha0 is a named development-runtime observation, not stable Office
compatibility or PowerPoint editing evidence. Empty/single/long cases have no
producer admission from this six-case run.

The matching [content-only controls](../../testing/smartart-native-candidate-producer-evidence.json)
retain their content graph but omit presentation nodes/edges. All six standard
controls have only the ordinary control text in extracted PDFs, even after
removing custT. This is a separate negative result: the tested runtime needs more
than a content forest for these layouts. The default content-only specimen is an
explicit negative control, not the native-output path. Presentation mode adds
the exact-layout mapping; it does not implement a general layout engine.

## Decision and next gates

Keep the modules private to research and retain `nativeAdmission: "untested"`.
Do not expose SmartArt creation/edit-save APIs yet. The next gates are independent
geometry/cache decisions, PowerPoint display and native SmartArt recognition,
node insertion/deletion/reordering, theme changes/re-layout and three native
rounds per exact layout. Family admission remains ADV-02.2.2; implementation,
managed IDs, source-preserving editing and preview belong to ADV-03.

Tests cover typed identities/transition reciprocity, presentation tree ownership,
Process detail mappings, assistants, Cycle closure, determinism, snapshotting,
invalid topology/XML, budgets/retry, Node/browser DOM agreement and three unrelated
text-save rounds. Synthetic component roots make structural tests redistributable;
they deliberately do not certify native layout execution.
