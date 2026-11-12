import { createPresentation, normalizePresentation, resolveTextFontSegments, type TextRunStyleInput } from "@pptkit/core";
import { resolveLayout } from "@pptkit/layout";
import { generatePptx } from "@pptkit/pptx-exporter";
import { renderPresentationToSvg } from "@pptkit/svg-renderer";

export async function createMultiscriptSlide() {
  const presentation = createPresentation();
  const style: TextRunStyleInput = {
    fontFamily: "Arial",
    eastAsianFontFamily: "Noto Sans CJK SC",
    complexScriptFontFamily: "Noto Sans Arabic",
    fontSize: 24,
  };
  presentation.addSlide().addElement({
    type: "text",
    box: { x: 40, y: 40, width: 600, height: 100 },
    content: [{ runs: [{ text: "Hello 中文 العربية", style }] }],
  });
  const normalized = normalizePresentation(presentation);
  const element = normalized.slides[0]!.elements[0]!;
  if (element.type !== "text") throw new Error("Expected text.");
  const run = element.content[0]!.runs[0]!;
  const segments = resolveTextFontSegments(run.text, run.style, normalized.theme, { maxSegments: 1000 });
  // A real measurement provider must verify loading and glyph coverage for
  // each selected segment; ordinary resolution retains authored geometry.
  const layout = resolveLayout(presentation);
  return { presentation, layout, segments, pptx: await generatePptx(presentation), svg: await renderPresentationToSvg(presentation) };
}
