# Generate and Preview a Deck with the Presentation Skill

The `pptkit-presentation` skill turns source material into a versioned deck session, opens an SVG review experience in a browser, and generates an editable PPTX only when the user requests a download. A Node project remains available as a fallback for unattended filesystem output and LibreOffice/PowerPoint-oriented validation.

## Install or update

In a skill-enabled agent such as Codex, ask it to install or update `pptkit-presentation` from `openHacking/pptkit`. If the agent does not support skill installation in chat, use:

```bash
npx skills add openHacking/pptkit --skill pptkit-presentation -g
```

Configure the deployed HTTPS review application as `PPTKIT_PREVIEW_URL`. The skill uses browser mode only when this URL and an in-app browser are available; otherwise it explains why it selected the Node fallback.

## Ask for a deck

Attach or name the relevant files and describe the audience and outcome:

```text
Use PPTKit to turn this quarterly report into an editable 10-slide presentation for our executive review.
```

The skill inspects available material, asks only for missing decisions one at a time, shows a theme preview, and proposes a slide-by-slide outline. It does not create artifacts until the user selects **Approve and generate**. **Change the plan** returns to the affected decision, while **Cancel** stops without artifacts.

## Browser-first review

After approval, browser mode creates `deck-brief.md`, `deck-session.json`, and `content/sources.json`.

The agent imports `DeckSessionV1` into the HTTPS review application. The page renders one standalone SVG per slide, shows blocking issues and warnings, stores the session and assets in IndexedDB, and keeps the review tab open. It does not upload deck data and does not generate PPTX bytes during preview.

Users can revise the deck in chat without losing their place: revisions retain stable slide IDs, increment the session revision, re-import the complete session, and report changed pages.

The explicit **Generate & download PPTX** action generates bytes in the browser, verifies ZIP/XML package structure, downloads `build-report.json`, and downloads the PPTX only when package checks pass. Browser SVG preview is a QA surface, not a pixel-identical PowerPoint renderer.

## Browser source support

The application accepts TXT/Markdown, PDF, DOCX, CSV/XLS/XLSX, PNG/JPEG/GIF, and SVG through browser `File`/`Blob` inputs. Evidence and binary assets remain in IndexedDB. PPTX parsing and template reuse are not supported.

Automated JSON imports may inline assets up to 5 MB each and 20 MB total. Larger assets require user file selection or the Node fallback.

## Node fallback

The skill initializes its isolated TypeScript starter when browser review is unavailable, assets cannot be transferred safely, unattended local output is required, or the user requests LibreOffice/PowerPoint-oriented rendering:

```bash
node skills/pptkit-presentation/scripts/init-project.mjs \
  --output /tmp/my-pptkit-deck \
  --title my-pptkit-deck \
  --theme clean-business
```

The Node adapter converts paths to the same portable asset contract, while deck authoring, validation, structure inspection, and byte-level package inspection come from `@pptkit/presentation-workflow`.

The fallback delivers `output/deck.pptx`, `output/build-report.json`, `deck-brief.md`, `src/deck-spec.ts`, `content/sources.json`, and optional rendered pages.
