# Reproducible import, export and layout debugging

This workflow separates authoring validation, package observations, static preview and native application evidence. It uses the installed `pptkit` CLI and creates new files; run it in a fresh directory and keep the original inputs. Reports can contain presentation text, paths and external relationship targets. Treat them as document content when sharing. The CLI does not fetch external relationships.

## Reproduce from authoring JSON

Create `deck.json` with the example from [CLI generation](../api/cli-generation.md), then run:

```sh
pptkit validate deck.json --export --json > validation.json
pptkit generate deck.json generated.pptx --json > generation.json
pptkit inspect generated.pptx --projection content --package-summary --json > inspection.json
pptkit inspect generated.pptx --projection content --capabilities --json > capabilities.json
pptkit preview generated.pptx --input-format pptx --slide 1 --json > preview.json
pptkit save generated.pptx copy.pptx --mode no-op --dry-run --json > plan.json
pptkit save generated.pptx copy.pptx --mode no-op --json > save.json
pptkit compare generated.pptx copy.pptx --json > comparison.json
```

Commands must each exit zero before using the next result. Shell redirection creates or truncates the named report even on failure; use fresh report names. The generated PPTX and copy use create-only publication and must not exist. A dry run does not check whether publication will be possible. `inputs` contains a source byte count and SHA-256: compare the captured hash across inspection, capability, preview, save and comparison reports to establish that they describe the same captured source. This does not prevent a later filesystem change.

Check Core diagnostics and export warnings in `validation.json` and `generation.json`. Validation with `--export` runs the exporter and structural parsing in memory; it does not prove native appearance. Inspection reports projected and unsupported source findings. Capability subjects are scoped observations, not permission to save; an absent observation is not support. `packageSummary` shows declared media sizes and relationship counts, without validating media content or proving that a part can be deleted. `preview.json` contains the static SVG and the complete import report; missing or approximated source content can affect it. The default comparison checks exact OPC payloads, not visual equivalence. A no-op source save is expected to retain original bytes, but a comparison result does not verify application editability.

For an existing PPTX, start at `inspect` using the original path. Preserve its exact hash, producer, application/version and reproduction steps; do not substitute a generated PPTX for the original when investigating a producer-specific failure. Use `--max-output-bytes` when necessary: a limit error means the report is unavailable, not that omitted findings are absent. The [CLI inspection](../api/cli-inspection.md), [capabilities](../api/import-capabilities.md), [preview](../api/cli-generation.md) and [save](../api/cli-save.md) contracts describe resource limits and profiles.

## Investigate layout

For an authoring document held in an application, use the checked [layout analysis example](../examples/layout-analysis.ts): call `resolveLayout` with `diagnostics.textOverflow` and `diagnostics.slideBounds`, supplying a `TextMeasurer` backed by the intended loaded fonts. Compare diagnostic codes, element IDs, source paths and measured bounds, then pass the same layout options to the exporter or renderer. Without a measurer, `text-measurement-unavailable` means overflow was not checked. The [layout contract](../api/layout.md#opt-in-measurement-and-diagnostics) details coordinate systems, exclusions and font limitations. The CLI preview does not provide a font-backed layout diagnostic switch; do not interpret its successful SVG as a text-fit check.

## Check native behavior separately

If application compatibility matters, open a copy in the named application and record its exact version, OS, repair prompts, visible differences, edit operation and reopened result. Record untested environments as such. Follow the [application compatibility matrix](application-compatibility.md) and link the source hash and retained evidence to the relevant task/fixture. Never infer native verification from an import report, successful SVG or package equality.

The [CLI workflow regression](../../packages/cli/test/commands.test.mjs) exercises the commands above against a repository-authored deck, checks source hashes, preview text and no-op comparison. It is automated package evidence, not a native application experiment.

For a preservation-first synthetic reproduction with source-relationship closure and explicit reduction/redaction boundaries, see the [worked example](minimal-reproduction.md).
