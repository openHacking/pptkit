import { createPresentation } from "@pptkit/core";
import { resolveLayout, type LayoutOptions } from "@pptkit/layout";
import { generatePptx } from "@pptkit/pptx-exporter";

export async function createPaginatedReport() {
  const document = createPresentation({ size: { width: 600, height: 400 } });
  const slide = document.addSlide({ id: "report", notes: "Notes remain on the source page." });
  slide.addElement({
    type: "table", id: "data", box: { x: 30, y: 60, width: 540, height: 1148 },
    columns: [180, 180, 180],
    rows: Array.from({ length: 41 }, (_, index) => ({
      height: 28,
      cells: [0, 1, 2].map((column) => ({ content: index === 0 ? `Column ${column + 1}` : `Row ${index}, cell ${column + 1}` })),
    })),
  });
  const layout: LayoutOptions = {
    tablePagination: [{ elementId: "data", area: { x: 30, y: 60, width: 540, height: 300 }, headerRows: 1, maxPages: 50 }],
  };
  const resolved = resolveLayout(document, layout);
  if (resolved.analysis?.diagnostics.length) return { diagnostics: resolved.analysis.diagnostics };
  return { fragments: resolved.pagination?.fragments, generated: await generatePptx(document, { layout }) };
}
