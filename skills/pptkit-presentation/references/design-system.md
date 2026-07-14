# Design system and slide roles

## Themes

### clean-business

Use a quiet white canvas, deep navy text, blue accents, thin dividers, and generous whitespace. Prefer Arial-compatible typography for predictable cross-application rendering. Use for executive reviews, product plans, project updates, and pitches.

### swiss-grid

Use a strict modular grid, black/white contrast, one red accent, square geometry, large numeric hierarchy, and compact metadata. Prefer Arial/Helvetica-compatible typography. Use for technical systems, product launches, data stories, and engineering content.

### editorial-story

Use warm paper, charcoal text, terracotta accents, serif display type, asymmetric image-led compositions, and restrained captions. Prefer Georgia plus Arial-compatible body text. Use for research, essays, cultural topics, and narrative reports.

## Ten roles

| Role | Use | Required plan fields |
| --- | --- | --- |
| `cover` | Open with title and framing | `title`, optional `subtitle` |
| `agenda` | Orient the audience | `items` |
| `section` | Mark a narrative transition | `title`, optional `message` |
| `statement` | Land one argument or quote | `message` |
| `image` | Pair evidence or product imagery with a point | `image`, optional `items` |
| `kpi` | Show 1–4 metrics | `kpis` |
| `comparison` | Compare two positions | `comparison` |
| `process` | Explain 2–6 ordered steps | `steps` |
| `table` | Present structured data | `table`, optional `chart` |
| `closing` | Close with action or takeaway | `message`, optional `items` |

## Density rules

- Keep one primary message per slide.
- Use at most six agenda/process items, four KPIs, two comparison columns, and eight visible table rows.
- Keep ordinary body text at or above 17 pt, captions at or above 11 pt, and titles at or above 28 pt.
- Prefer another slide over shrinking text or layering unrelated elements.
- Use a chart only when comparison or change is clearer visually than as a table.
- Use no more than two chart series in v1; label values directly when practical.

## Images

- Copy PNG, JPEG, GIF, or SVG images into `assets/` and reference relative paths from `deck-spec.ts`. Convert WebP/HEIC assets before authoring because the v1 exporter does not preserve those encodings safely.
- Provide width and height from `content/sources.json` so PPTKit can resolve `contain` and `cover` deterministically.
- Use `cover` for photographic crops and `contain` for screenshots, diagrams, and UI evidence.
- Preserve sensitive or text-heavy screenshots; do not redraw them without explicit permission.
