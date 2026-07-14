---
name: pptkit-presentation
description: Create polished, structured, editable PowerPoint presentations with PPTKit from a topic, pasted text, Markdown, TXT, PDF, DOCX, CSV, XLSX, or image assets. Use when the user asks to make, generate, design, revise, or export a PPT, PPTX, slide deck, presentation, pitch deck, report deck, or training deck, including requests such as “create a presentation”, “turn this report into slides”, or “create an editable PowerPoint”.
---

# PPTKit Presentation

Create a local TypeScript project whose source of truth is a structured deck specification and whose final output is an editable `.pptx` generated through PPTKit public APIs.

## Run the workflow

1. Inspect the request and every supplied source before asking questions. Read [workflow.md](references/workflow.md).
2. Ask only for decisions that cannot be inferred. Ask them one at a time in this order: purpose and audience, theme, then page count and asset strategy. Use the host's native question/form tool (in Codex, `request_user_input`) whenever available; each question must offer a recommended, mutually exclusive set of choices and free-form input. Otherwise use the numbered fallback in [workflow.md](references/workflow.md), then stop and wait for the reply.
3. Show the three style previews in `assets/previews/` unless the user already chose a theme. Recommend exactly one theme.
4. Build the normalized brief and slide-by-slide outline only after collecting the needed decisions. Keep the detailed outline separate from the short decision summary; use the host's expandable detail section when it provides one, but never rely on that section for the confirmation control.
5. Require the user to choose exactly one confirmation outcome: **Approve and generate**, **Change the plan**, or **Cancel**. Do not initialize a project, install dependencies, copy sources, or generate a PPTX before **Approve and generate**. On **Change the plan**, return to the affected decision and refresh the outline. On **Cancel**, summarize the collected brief without creating project artifacts. Skip this gate only when the user supplied a complete specification and explicitly asked to generate without confirmation.
6. Initialize a project after approval:

   ```bash
   node <SKILL_DIR>/scripts/init-project.mjs --output <PROJECT_DIR> --title <SLUG> --theme <THEME_ID>
   ```

   Use `--no-install` only when dependencies are already unavailable or when preparing files without executing them. Never modify the skill's bundled starter in place.
7. Copy source files into `<PROJECT_DIR>/sources/`, then run `npm run extract -- <paths...>`. Read the generated `content/sources.json` and use source IDs in the outline.
8. Write `deck-brief.md` and edit only `src/deck-spec.ts` for normal deck creation. Read [design-system.md](references/design-system.md) before choosing slide roles or density. Read [pptkit-api.md](references/pptkit-api.md) only when extending the bundled runtime.
9. Run `npm run build`. Treat validation errors, out-of-bounds elements, missing required assets, malformed package parts, and unexpected exporter warnings as failures. Fix and rerun.
10. Run `npm run render`. Rendering is optional: if LibreOffice or Poppler is unavailable, keep the structural result and explicitly request manual review in PowerPoint or LibreOffice. Read [quality.md](references/quality.md).
11. Deliver the `.pptx`, `deck-brief.md`, `src/deck-spec.ts`, `content/sources.json`, `output/build-report.json`, and optional `output/rendered/` artifacts. Mention every remaining warning.

## Keep these contracts

- Keep `DeckBrief`, `SlidePlan`, `ThemeId`, and `BuildReport` compatible with `src/contracts.ts`.
- Use one of `clean-business`, `swiss-grid`, or `editorial-story`.
- Use the ten supported roles: `cover`, `agenda`, `section`, `statement`, `image`, `kpi`, `comparison`, `process`, `table`, and `closing`.
- Use native PPTKit text, shapes, connectors, images, and tables. Use editable shapes/connectors for bar and line charts; do not claim that they are native data-bound PowerPoint charts.
- Keep source material local. Copy images into the project instead of relying on temporary paths.
- Use only PPTKit public exports. Do not import `dist` files or private implementation paths.
- Do not claim template reuse, PPTX parsing, animation, SmartArt, audio/video, or browser editing.
- Do not copy templates, code, or assets from other presentation skills.

## Control quality

- Prefer a clear narrative and one message per slide over filling every available region.
- Preserve citations and source IDs in `sourceRefs`; never invent figures or claims.
- Split content instead of shrinking body text below the theme minimum.
- Use the smallest sufficient visual vocabulary and reuse theme tokens.
- Run the complete build after every material revision; a generated file is not proof of a correct deck.
- Perform a final adversarial review against the failure list in [quality.md](references/quality.md).

## Compatibility

Read [compatibility.md](references/compatibility.md) before changing dependency versions or publishing the skill. Old PPTX files may be treated as reference material only; the v1 workflow has no PPTX parser or template-fill contract.
