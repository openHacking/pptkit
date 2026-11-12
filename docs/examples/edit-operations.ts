import { createPresentation, type PresentationEditOperation } from "@pptkit/core";

const deck = createPresentation({ id: "deck" });
deck.addSlide({ id: "slide-1", notes: [{ runs: [{ text: "Draft notes" }] }], elements: [{ id: "title", type: "text", content: [{ runs: [{ text: "Draft" }] }] }] });
deck.addSlide({ id: "slide-2", elements: [{ id: "chart", type: "chart", chartType: "bar", box: { x: 20, y: 20, width: 200, height: 100 }, categories: ["Q1"], series: [{ name: "Revenue", values: [1] }] }] });
deck.slides[0]!.setReadingOrder([{ elementIds: ["title"] }]);
const operation: PresentationEditOperation = {
  schemaVersion: "pptkit-edit-v1", documentId: "deck",
  changes: [
    { type: "slide-hidden", slideId: "slide-1", expected: false, value: true },
    { type: "first-slide-number", expected: null, value: 0 },
    { type: "slide-order", expected: ["slide-1", "slide-2"], value: ["slide-2", "slide-1"] },
    { type: "text-run", slideId: "slide-1", elementId: "title", paragraphIndex: 0, runIndex: 0, expected: "Draft", value: "Final" },
    { type: "element-name", slideId: "slide-1", elementId: "title", expected: null, value: "Title" },
    { type: "element-accessibility", slideId: "slide-1", elementId: "title", expected: null, value: { description: "Presentation title", decorative: false } },
    { type: "element-transform", slideId: "slide-1", elementId: "title", expected: null, value: { rotation: 15, flipH: false, flipV: null } },
    { type: "element-action", slideId: "slide-1", elementId: "title", expected: null, value: { type: "slide", slideId: "slide-2", tooltip: "Next" } },
    { type: "text-run-action", slideId: "slide-1", elementId: "title", paragraphIndex: 0, runIndex: 0, expected: null, value: { type: "url", url: "https://example.org", tooltip: null } },
    { type: "slide-reading-order", slideId: "slide-1", expected: [{ elementIds: ["title"] }], value: null },
    { type: "notes-run", slideId: "slide-1", paragraphIndex: 0, runIndex: 0, expected: "Draft notes", value: "Final notes" },
    { type: "chart-point", slideId: "slide-2", elementId: "chart", seriesIndex: 0, pointIndex: 0, coordinate: "values", expected: 1, value: "-0" },
    { type: "chart-category", slideId: "slide-2", elementId: "chart", categoryIndex: 0, expected: "Q1", value: "First quarter" },
    { type: "chart-series-name", slideId: "slide-2", elementId: "chart", seriesIndex: 0, expected: "Revenue", value: "Net revenue" },
  ],
};
const undo = deck.applyEditOperation(operation).inverse;
const redo = deck.applyEditOperation(undo).inverse;
deck.applyEditOperation(redo);
