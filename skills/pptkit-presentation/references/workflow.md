# Guided workflow

## Intake

Derive everything possible from the user's prompt and files first. Ask only for missing decisions that change the result. Ask one decision at a time, in this order:

1. **Purpose:** audience, delivery setting, desired outcome, and language.
2. **Design:** preferred theme, brand constraints, format, and non-negotiable content.
3. **Scope and material:** duration/page range, authoritative sources, image policy, confidentiality, and citations.

## Interaction capability

Prefer the host's native user-input capability. Use single-select, multi-select, or text input as the decision needs; include a recommended choice, mutually exclusive choices where applicable, and free-form input. In Codex, call `request_user_input`; on another agent, call its equivalent native question/form tool when it exists.

If the host has no native user-input capability, use this fallback exactly, then stop and wait for a reply:

```text
Choose one option for <decision>:
1. <recommended option> — <effect>
2. <alternative> — <effect>
3. Other — reply with your preference
```

Do not claim that this skill creates a custom in-chat plugin form. It can request the host's native control, but the host owns its UI.

## Theme selection

Show these local previews and recommend one:

- `assets/previews/clean-business.svg` — restrained product, strategy, status, and executive decks.
- `assets/previews/swiss-grid.svg` — technical, data-led, engineering, and launch narratives.
- `assets/previews/editorial-story.svg` — research, cultural, thought-leadership, and story-led decks.

Respect an explicit user choice. Never mix themes within one deck.

## Confirmation gate

After the outline is ready, show a short decision summary that remains visible:

- working title and expected slide count
- selected theme and image strategy
- material constraints or known gaps
- outline overview

Put the full per-slide plan (role, title, message, visual, and source IDs) after the summary. Use an expandable detail section when the host provides one, but do not rely on it for the confirmation control. Then request exactly one outcome through the host's native control, or the numbered fallback:

1. **Approve and generate** — create the deck session and continue to preview.
2. **Change the plan** — ask again only for the affected decision, refresh the outline, and present confirmation again.
3. **Cancel** — summarize the collected brief and do not create project artifacts.

Treat every outcome except **Approve and generate** as a stop. Do not create session/project artifacts, open a preview, install dependencies, copy sources, or generate PPTX bytes before approval. Skip the gate only if the user supplied a complete specification and explicitly requested generation without confirmation.

## Project artifacts

Use this browser-first project shape:

```text
project/
├── deck-brief.md
├── content/sources.json
└── deck-session.json
```

The browser source of truth is `deck-session.json`; extracted content is evidence, not slide copy. The Node fallback retains its isolated TypeScript project shape from `node-workflow.md`.

## Iteration

Translate feedback into changes to the brief or slide plan, then rebuild the whole deck. Keep stable slide IDs when the slide's semantic identity is unchanged. Report which pages changed and whether warnings remain.
