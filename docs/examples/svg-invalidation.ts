import type { LayoutResult } from "@pptkit/layout";
import { captureSvgPreviewSnapshot, planSvgPreviewInvalidation, renderLayoutToSvg, renderIncrementalSvg, type IncrementalSvgCache } from "@pptkit/svg-renderer";

/** The caller retains the previous snapshot and publishes only if this input is still current. */
export async function renderChangedPages(previous: LayoutResult, next: LayoutResult, environmentRevision: string) {
  const options = { environmentRevision };
  const plan = planSvgPreviewInvalidation(captureSvgPreviewSnapshot(previous, options), captureSvgPreviewSnapshot(next, options));
  const changed = await renderLayoutToSvg(next, { slideIndices: plan.render.map(item => item.index) });
  return { plan, changed };
}

/** Keep the returned cache only if the caller's document revision is still current. */
export function renderPreview(layout: LayoutResult, environmentRevision: string, cache?: IncrementalSvgCache) {
  return renderIncrementalSvg(layout, { environmentRevision }, cache);
}
