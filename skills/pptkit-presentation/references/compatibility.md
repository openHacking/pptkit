# Compatibility and release

## Runtime matrix

| Component | v1 contract |
| --- | --- |
| Node.js | `>=20` |
| `@pptkit/core` | exact `0.1.1` in the starter |
| `@pptkit/pptx-exporter` | exact `0.1.1` in the starter |
| TypeScript | exact version in the starter manifest |
| PowerPoint/LibreOffice | manual compatibility review required for public release |

Keep PPTKit package versions aligned. Update the starter only after the workspace build, typecheck, lint, tests, package dry-run, and representative deck fixtures pass together.

## Installation

Install the skill with:

```bash
npx skills add openHacking/pptkit --skill pptkit-presentation -g
```

Skill installation does not install project dependencies globally. `init-project.mjs` creates an isolated project and installs its pinned dependencies on first use.

## Publishing boundary

The public installation path requires the matching PPTKit versions to exist on npm. Do not publish the skill as generally available while its pinned PPTKit packages are unavailable. npm publication requires maintainer credentials and explicit release approval.

## Licensing

The skill implementation and visual assets must be original PPTKit work. External presentation skills may inform workflow research only; do not copy their source, templates, previews, or proprietary exporters.
