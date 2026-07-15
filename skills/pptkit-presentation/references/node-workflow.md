# Node fallback workflow

Use this workflow after approval when the browser workflow is unavailable or unsuitable. State the fallback reason.

1. Initialize an isolated project:

   ```bash
   node <SKILL_DIR>/scripts/init-project.mjs --output <PROJECT_DIR> --title <SLUG> --theme <THEME_ID>
   ```

   Use `--no-install` only when dependencies are unavailable or when preparing files without execution. Never modify the bundled starter.
2. Copy source files into `<PROJECT_DIR>/sources/`, run `npm run extract -- <paths...>`, and read `content/sources.json`.
3. Write `deck-brief.md` and edit `src/deck-spec.ts`. Reference copied images by `assetId`, relative to the project's `assets/` directory.
4. Run `npm run build`. The shared workflow runtime validates the deck, inspects layout, generates bytes through the Node asset adapter, writes `output/deck.pptx`, and checks the package from `Uint8Array`.
5. Run `npm run render` when LibreOffice/Poppler review is required. Missing render tools produce a reported skip, not a false success.
6. Deliver `output/deck.pptx`, `output/build-report.json`, `deck-brief.md`, `src/deck-spec.ts`, `content/sources.json`, and optional `output/rendered/`.

Keep local assets inside `assets/`; the Node adapter is the only layer allowed to resolve filesystem paths.
