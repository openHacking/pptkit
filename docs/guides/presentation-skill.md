# Generate a Deck with the Presentation Skill

The `pptkit-presentation` skill packages PPTKit's developer APIs as a guided workflow for local coding agents. It produces an editable `.pptx`, the TypeScript deck specification, extracted source evidence, and a machine-readable build report.

## Install or update

In a skill-enabled agent such as Codex, you do not need a terminal. Send one of these
sentences in chat:

> Install the `pptkit-presentation` skill from the GitHub repository `openHacking/pptkit`.

> Update my installed `pptkit-presentation` skill to the latest version from the GitHub repository `openHacking/pptkit`.

The agent should confirm the installation or update before you use the skill. Updating
reinstalls the skill files; it does not alter deck projects that were already generated.
Use the same update sentence after a new version is published.

If your agent does not offer skill installation in chat, use its documented installer.
For the command-line installer, run:

```bash
npx skills add openHacking/pptkit --skill pptkit-presentation -g
```

The open Agent Skills format is supported by Codex, Claude Code, Cursor, and other compatible local agents. Skill installation is separate from project dependencies: the skill creates an isolated Node.js project and installs its pinned packages when the first deck starts.

For repository-local evaluation before npm publication, initialize the starter directly:

```bash
node skills/pptkit-presentation/scripts/init-project.mjs \
  --output /tmp/my-pptkit-deck \
  --title my-pptkit-deck \
  --theme clean-business \
  --no-install
```

Use `--no-install` only for workspace evaluation. A normal installed-skill run omits it.

## Ask for a deck

Start with a plain request and attach or name the relevant files:

```text
Use PPTKit to turn this quarterly report into an editable 10-slide PPTX for our executive review.
```

The skill inspects the material, then asks only for decisions it cannot infer, one at a time: purpose and audience, theme, then scope and asset strategy. It uses the host's native question control when one is available (Codex uses its native option card); other agents use their equivalent control. A host without that capability receives a numbered text choice and the skill waits for a reply.

Before any project is created, the skill shows a short, always-visible decision summary and keeps the full slide-by-slide outline separate; it uses an expandable detail section when the host provides one. The user must choose **Approve and generate**, **Change the plan**, or **Cancel**. Only approval permits project initialization and PPTX generation; change reopens only the affected decision, and cancel leaves no project artifacts. This confirmation is skipped only for a complete brief that explicitly requests generation without confirmation. The skill cannot add a custom ChatCut-style form to hosts that do not provide one.

The bundled themes are:

- `clean-business` for executive, product, strategy, and status decks
- `swiss-grid` for technical, data-led, engineering, and launch decks
- `editorial-story` for research, cultural, and narrative decks

## Supported material

The deterministic extractor accepts Markdown, TXT, PDF, DOCX, CSV, XLSX, and PNG/JPEG/GIF/SVG images. It writes normalized records with stable source IDs to `content/sources.json` and copies images into the local project. Convert WebP or HEIC files before use. Old PPTX files are reference material only; v1 does not parse, fill, or reuse their masters.

## Build and review

The agent edits `src/deck-spec.ts`, then runs:

```bash
npm run build
npm run render
```

The build validates the presentation, checks slide boundaries and risky overlaps, exports through `@pptkit/pptx-exporter/node`, and verifies required ZIP/XML parts. Any exporter warning is recorded and fails the build until fixed.

Rendering is optional. When LibreOffice, Poppler, and ImageMagick are available, the render command creates slide PNGs and a contact sheet. Otherwise the structurally valid PPTX remains available and must be opened manually in PowerPoint or LibreOffice.

## Output

Deliver these artifacts together:

- `output/deck.pptx`
- `output/build-report.json`
- `deck-brief.md`
- `src/deck-spec.ts`
- `content/sources.json`
- `output/rendered/` when visual rendering is available

The v1 skill uses native editable PPTKit objects. Its bar and line charts are editable shapes and connectors, not native data-bound PowerPoint chart objects.
