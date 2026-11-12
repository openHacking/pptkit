import { pixelPng } from "../../tests/fixtures/presentations.mjs";

export const workloads = { small: 10, medium: 50, large: 200 };

export function createBenchmarkDeck(createPresentation, slideCount) {
  if (!Number.isSafeInteger(slideCount) || slideCount < 1) throw new Error("slideCount must be a positive safe integer");
  const deck = createPresentation({ id: "benchmark", metadata: { title: "PPTKit mixed workload", author: "PPTKit" }, theme: { fonts: { heading: "Arial", body: "Arial" } } });
  deck.registerAsset({ id: "image", kind: "image", source: { type: "url", value: pixelPng }, mimeType: "image/png", width: 8, height: 4 });
  for (let index = 0; index < slideCount; index++) {
    const slide = deck.addSlide({ id: `slide-${index}`, notes: `Speaker notes ${index}` });
    slide.addElement({ type: "text", id: `title-${index}`, content: `Quarterly report ${index + 1}`, box: { x: 40, y: 24, width: 880, height: 60 } });
    slide.addElement({ type: "text", id: `body-${index}`, content: [
      { runs: [{ text: "Structured ", style: { bold: true } }, { text: "editable content & reliable output" }] },
      { runs: [{ text: "A second paragraph for deterministic normalization and preview." }] },
    ], box: { x: 40, y: 100, width: 540, height: 120 } });
    slide.addElement({ type: "image", id: `image-${index}`, assetId: "image", fit: "cover", box: { x: 650, y: 100, width: 220, height: 110 } });
    slide.addElement({ type: "group", id: `group-${index}`, box: { x: 650, y: 260, width: 220, height: 160 }, coordinateSize: { width: 220, height: 160 }, children: [
      { type: "shape", id: `shape-${index}`, shape: "roundRect", box: { x: 0, y: 0, width: 220, height: 160 }, text: { content: "Summary" } },
    ] });
    if (index % 2 === 0) slide.addElement({ type: "chart", id: `chart-${index}`, chartType: "bar", categories: ["A", "B", "C", "D"], series: [{ name: "Actual", values: [3, 7, 5, 9] }, { name: "Plan", values: [4, 6, 6, 8] }], box: { x: 40, y: 260, width: 540, height: 240 } });
    else slide.addElement({ type: "table", id: `table-${index}`, box: { x: 40, y: 260, width: 540, height: 240 }, columns: [180, 180, 180], rows: Array.from({ length: 4 }, (_, row) => ({ height: 60, cells: ["Metric", `Value ${row}`, `Target ${row}`].map((content) => ({ content })) })) });
  }
  return deck;
}

export function summarize(samples) {
  if (!samples.length || samples.some((value) => !Number.isFinite(value))) throw new Error("Expected nonempty finite samples");
  const sorted = [...samples].sort((a, b) => a - b);
  const mean = samples.reduce((sum, value) => sum + value, 0) / samples.length;
  const variance = samples.reduce((sum, value) => sum + (value - mean) ** 2, 0) / samples.length;
  const midpoint = Math.floor(sorted.length / 2);
  return {
    samples, min: sorted[0], max: sorted.at(-1), mean,
    median: sorted.length % 2 ? sorted[midpoint] : (sorted[midpoint - 1] + sorted[midpoint]) / 2,
    p95: sorted[Math.ceil(sorted.length * 0.95) - 1],
    standardDeviation: Math.sqrt(variance), coefficientOfVariation: mean === 0 ? 0 : Math.sqrt(variance) / mean,
  };
}
