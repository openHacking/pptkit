# Guided workflow

## Intake

Derive everything possible from the user's prompt and files first. Ask only for missing decisions that change the result. Group the intake into at most three questions:

1. **Purpose:** audience, delivery setting, desired outcome, language, and duration/page range.
2. **Material:** authoritative sources, required claims/data, image policy, confidentiality, and citations.
3. **Design:** preferred theme, brand constraints, format, and non-negotiable content.

When the host offers structured questions, use it. Otherwise ask the same fields in concise prose. Do not pretend that a custom in-chat plugin form exists.

## Theme selection

Show these local previews and recommend one:

- `assets/previews/clean-business.svg` — restrained product, strategy, status, and executive decks.
- `assets/previews/swiss-grid.svg` — technical, data-led, engineering, and launch narratives.
- `assets/previews/editorial-story.svg` — research, cultural, thought-leadership, and story-led decks.

Respect an explicit user choice. Never mix themes within one deck.

## Confirmation gate

Confirm these fields together:

- working title, audience, purpose, language, and expected slide count
- selected theme and image strategy
- hard constraints and known gaps
- one line per slide: role, title, message, visual, and source IDs

Wait for explicit confirmation. Skip the wait only if the user supplied a complete specification and explicitly requested uninterrupted generation.

## Project artifacts

Use this project shape:

```text
project/
├── deck-brief.md
├── sources/
├── assets/
├── content/sources.json
├── src/deck-spec.ts
└── output/
    ├── deck.pptx
    ├── build-report.json
    └── rendered/             # optional
```

The editable source of truth is `src/deck-spec.ts`; extracted content is evidence, not slide copy. Summarize and structure it deliberately.

## Iteration

Translate feedback into changes to the brief or slide plan, then rebuild the whole deck. Keep stable slide IDs when the slide's semantic identity is unchanged. Report which pages changed and whether warnings remain.
