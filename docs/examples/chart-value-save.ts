import { importPptx } from "@pptkit/pptx-parser";

export async function saveExistingChartValue(
  source: Uint8Array, seriesIndex: number, pointIndex: number, value: number,
) {
  const session = await importPptx(source, {
    projection: "content",
    chartEditSaveProfile: "existing-values-v1",
  });
  try {
    const slide = session.document.slides.find(s => s.elements.some(e => e.type === "chart"));
    const chart = slide?.elements.find(e => e.type === "chart");
    if (!slide || !chart || chart.type !== "chart" || chart.chartType === "scatter" || chart.chartType === "bubble") throw new Error("No mapped chart.");
    const binding = session.report.chartValues?.find(b => b.elementId === chart.id &&
      b.slideId === slide.id && b.seriesIndex === seriesIndex && b.pointIndex === pointIndex);
    if (binding?.status !== "bound") throw new Error(binding?.reason ?? "No numeric source binding.");
    const updated = structuredClone(chart), series = updated.series[seriesIndex];
    if (!Number.isFinite(value) || !series || pointIndex < 0 || pointIndex >= series.values.length) {
      throw new Error("Expected a finite value at an existing point.");
    }
    series.values[pointIndex] = value;
    const index = slide.elements.indexOf(chart);
    slide.removeElement(chart.id);
    slide.insertElement(index, updated);
    const prepared = await session.prepareSave({ mode: "edit" });
    if (prepared.status !== "ready") throw new Error(prepared.report.diagnostics.map(d => d.message).join("\n"));
    return await session.save(prepared.plan);
  } finally { session.dispose(); }
}
