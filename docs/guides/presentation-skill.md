# Generate and Preview a Deck with the Presentation Skill

The `pptkit-presentation` skill turns source material into a versioned deck session, opens an SVG review experience in a browser, and generates an editable PPTX only when the user requests a download. A Node project remains available as a fallback for unattended filesystem output and LibreOffice/PowerPoint-oriented validation.

## Install or update

In a skill-enabled agent such as Codex, install the skill by sending:

```text
Install the `pptkit-presentation` skill from the GitHub repository `openHacking/pptkit`.
```

To update the installed skill from the same repository, send:

```text
Update the `pptkit-presentation` skill from the GitHub repository `openHacking/pptkit`.
```

If the agent does not support skill installation in chat, install it globally with:

```bash
npx skills add openHacking/pptkit --skill pptkit-presentation -g
```

Update installed skills with:

```bash
npx skills update
```

The installed skill uses `https://openhacking.github.io/pptkit/` as its default HTTPS review application, so customers do not need to configure a preview URL. For a private deployment, staging environment, or local development, supply a URL for the current task or set `PPTKIT_PREVIEW_URL`; either overrides the official default. In Codex, the skill discovers the in-app Browser controls and `node_repl js`, initializes the `iab` browser, and attempts to open the resolved URL before declaring browser review unavailable. It does not treat an abbreviated initial tool list as evidence that no browser exists. The skill selects the Node fallback only after a real browser setup/navigation failure or when another documented fallback condition applies, and it explains the reason.

## Ask for a deck

Attach or name the relevant files and describe the audience and outcome:

```text
Use PPTKit to turn this quarterly report into an editable 10-slide presentation for our executive review.
```

The skill inspects available material, asks only for missing decisions one at a time, shows a theme preview, and proposes a slide-by-slide outline with a composition intent and density check for every page. It does not create artifacts until the user selects **Approve and generate**. **Change the plan** returns to the affected decision, while **Cancel** stops without artifacts.

The generated deck keeps source IDs and filenames out of visible slide copy. `sourceRefs` remain available as provenance and are written to speaker notes; visible citations are included only when the user requests a human-readable citation treatment.

## Browser-first review

After approval, browser mode creates `deck-brief.md`, `deck-session.json`, and `content/sources.json`.

The agent transfers `DeckSessionV1` JSON bytes and every referenced asset through the same resumable `pptkit-transfer-v1` protocol. The page renders one standalone SVG per slide, shows blocking issues and warnings, stores the session and assets in IndexedDB, and keeps the review tab open. It does not upload deck data and does not generate PPTX bytes during preview.

Users can revise the deck in chat without losing their place: revisions retain stable slide IDs, increment the session revision, re-import the complete session, and report changed pages.

All runtimes use the same theme-specific authoring recipes. Clean Business emphasizes information axes and rules, Swiss Grid uses modular asymmetry and numeric anchors, and Editorial Story uses narrow measures and narrative image/text compositions.

After preview, the explicit **Generate & download PPTX** action—clicked by the user or triggered by the agent after an explicit user request—generates bytes in the browser, verifies ZIP/XML package structure, downloads `build-report.json`, and downloads the PPTX only when package checks pass. Browser SVG preview is a QA surface, not a pixel-identical PowerPoint renderer.

## Browser source support

The agent inspects TXT/Markdown, PDF, DOCX, CSV/XLS/XLSX, PNG/JPEG/GIF, and SVG sources with host-native tools, then sends normalized session evidence and supported assets through the chunk protocol. Evidence and binary assets remain in IndexedDB. PPTX parsing and template reuse are not supported.

Sessions never contain `dataUrl` assets. File size alone does not select Node; browser storage quota and verified transfer results determine whether Browser mode can continue.

The preview exposes a hidden, read-only `[data-testid="pptkit-preview-bridge"]` DOM integration surface. Its JSON contains the protocol name, maximum chunk size, API availability measured by the page itself, and resumable transfer progress. Codex Browser reads this DOM contract because its isolated read-only evaluation sandbox is intentionally not the page's native global context. The legacy `window.__pptkitPreviewBridge` remains available for ordinary browser automation. State changes are accepted only after opening the progressive `[data-testid="pptkit-transfer-toggle"]` surface and using the stable `pptkit-transfer-input` and `pptkit-transfer-submit` DOM controls. This keeps browser automation auditable without leaving a human-oriented payload form in the finished review UI.

## Node fallback

The skill initializes its isolated TypeScript starter when browser review is unavailable, the unified transfer protocol actually fails, unattended local output is required, or the user requests LibreOffice/PowerPoint-oriented rendering. Runtime routing is a guarded decision: the initializer refuses to create a project unless the caller supplies a valid reason, matching browser-check status and step, and concrete evidence.

```bash
node skills/pptkit-presentation/scripts/init-project.mjs \
  --output /tmp/my-pptkit-deck \
  --title my-pptkit-deck \
  --theme clean-business \
  --fallback-reason strict-office-rendering \
  --browser-check not-required \
  --browser-step user-requirement \
  --fallback-evidence "The user explicitly requested LibreOffice rendering"
```

The Node adapter converts paths to the same portable asset contract, while deck authoring, validation, structure inspection, and byte-level package inspection come from `@pptkit/presentation-workflow`.

The initializer records the accepted routing receipt in `runtime-decision.json`. The fallback also delivers `output/deck.pptx`, `output/build-report.json`, `deck-brief.md`, `src/deck-spec.ts`, `content/sources.json`, and optional rendered pages.
