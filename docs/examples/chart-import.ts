import { importPptx } from "@pptkit/pptx-parser";
import { generatePptx } from "@pptkit/pptx-exporter";

export async function exportProjectedChartEdit(source: Uint8Array) {
  const session = await importPptx(source, {
    projection: "content",
    sessionLimits: { maxChartDataPoints: 100000, maxExpandedChartWorkbookBytes: 16 * 1024 * 1024 },
  });
  try {
    const slide = session.document.slides.find(s => s.elements.some(e => e.type === "chart"));
    if (!slide) throw new Error("No chart in the supported projection.");
    const index = slide.elements.findIndex(e => e.type === "chart");
    const chart = slide.elements[index]!;
    if (chart.type !== "chart" || chart.chartType === "scatter" || chart.chartType === "bubble" || chart.chartType === "combo") throw new Error("Expected a mapped chart.");
    const changed = { ...chart, series: chart.series.map(s => ({ ...s, values: s.values.map(v => v === null ? null : v * 2) })) };
    slide.removeElement(chart.id);
    slide.insertElement(index, changed);
    // Review importReport before adopting fresh output: unsupported source
    // properties are omitted. This is a new projection document.
    const fresh = await generatePptx(session.document);
    const preservation = await session.prepareSave({ mode: "edit" });
    return { fresh, importReport: session.report, preservation };
  } finally {
    session.dispose();
  }
}
