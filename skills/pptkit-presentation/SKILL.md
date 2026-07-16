---
name: pptkit-presentation
description: Create polished, structured, editable PowerPoint presentations with PPTKit from a topic, pasted text, Markdown, TXT, PDF, DOCX, CSV, XLSX, or image assets. Use when the user asks to make, generate, design, revise, preview, or export a PPT, PPTX, slide deck, presentation, pitch deck, report deck, or training deck, including requests such as “create a presentation”, “turn this report into slides”, or “create an editable PowerPoint”.
---

# PPTKit Presentation

Create a structured deck session, preview it in the browser, and generate an editable `.pptx` only after the user explicitly chooses export. Prefer the browser workflow; keep the Node project workflow as a compatibility and strict-rendering fallback.

## Run the workflow

1. Inspect the request and every supplied source before asking questions. Read [workflow.md](references/workflow.md).
2. Ask only for decisions that cannot be inferred. Ask them one at a time in this order: purpose and audience, theme, then page count and asset strategy. Use the host's native question/form tool (in Codex, `request_user_input`) whenever available.
3. Show the three style previews in `assets/previews/` unless the user already chose a theme. Recommend exactly one theme.
4. Build the normalized brief and slide-by-slide outline. Record each slide's role, composition intent, density, visual evidence, and source IDs. Keep the detailed outline separate from the short decision summary.
5. Require exactly one confirmation outcome: **Approve and generate**, **Change the plan**, or **Cancel**. Do not create artifacts, open a preview, install dependencies, or generate PPTX bytes before approval. Skip this gate only for a complete specification that explicitly requests generation without confirmation.
6. After approval, choose the runtime:
   - Use [browser-workflow.md](references/browser-workflow.md) when an in-app browser and the official or configured HTTPS preview URL are available, assets fit the transfer limits, and the user did not require unattended local output or Office/LibreOffice rendering.
   - Use [node-workflow.md](references/node-workflow.md) otherwise, and state the fallback reason.
7. Read [design-system.md](references/design-system.md) before authoring. Use `deck-session.json` as the browser source of truth and stable slide IDs across revisions.
8. Treat validation errors, missing required assets, out-of-bounds elements, risky overlaps, malformed packages, and unexpected exporter warnings as failures. Read [quality.md](references/quality.md).
9. Deliver the browser preview first and do not download automatically. After preview, export only when the user clicks **Generate & download PPTX** or explicitly asks the agent to export/download it. Mention every remaining warning and the SVG-versus-Office fidelity boundary.

## Keep these contracts

- Use `DeckSessionV1` with `schemaVersion: 1`; keep `DeckBrief`, `SlidePlan`, `ThemeId`, and `BuildReport` compatible with `@pptkit/presentation-workflow`.
- Use `assetId` references in `ImagePlan`; never leak temporary filesystem paths into the browser-neutral deck spec.
- Use one of `clean-business`, `swiss-grid`, or `editorial-story` and the ten supported slide roles.
- Use native PPTKit text, shapes, connectors, images, and tables. Editable shape-based charts are not native data-bound PowerPoint charts.
- Keep source material local. Browser sessions use IndexedDB and `data:`/`blob:` assets; Node projects copy assets into `assets/`.
- Use only PPTKit public exports. Do not import `dist` files or private implementation paths.
- Do not claim PPTX parsing, template fill, animation, SmartArt, audio/video, browser editing, cross-device preview links, or pixel-identical PowerPoint preview.
- Do not copy templates, code, or assets from other presentation skills.

## Control quality

- Prefer one message per slide; preserve citations and source IDs in `sourceRefs`, which are provenance metadata and are not visible slide footers.
- Never place internal source IDs, input filenames, local paths, template/style names, or workflow instructions in visible slide copy. If a visible citation is explicitly required, author a human-readable citation as content and keep the internal ID in `sourceRefs`.
- Split content instead of shrinking below the theme minimum. Treat 18–22 pt as ordinary body copy, 15–18 pt as detail/table copy, and 9–11 pt as metadata only.
- Use theme-specific compositions and vary narrative rhythm. Do not solve empty space with filler, decorative numbering, repeated rounded cards, or arbitrary icons.
- Rebuild and re-preview after each material revision; keep the current slide selected by stable slide ID.
- Treat SVG renderer warnings as review evidence, not as proof of Office fidelity.
- Perform the final adversarial review in [quality.md](references/quality.md).

## Compatibility

Read [compatibility.md](references/compatibility.md) before changing dependency versions, the preview URL, or publishing the skill. Old PPTX files remain reference material only.
