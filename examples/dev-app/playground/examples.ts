import { createPresentation } from "@pptkit/core";
import { createFixture, fixtureCatalog } from "../../../tests/fixtures/presentations.mjs";
import { getExampleById } from "../src/example-registry";
import { createExamplePresentation } from "../src/presentation-builder";
import { parseExampleSource } from "../src/source-parser";
import nativeSampleUrl from "../../../tests/fixtures/corpus-text-powerpoint-interactive.pptx?url";

const fullDeck = getExampleById("export-full-feature-deck");
if (!fullDeck) throw new Error("The full feature deck is missing from the workbench registry.");

export const defaultExample = {
  id: fullDeck.id,
  title: "Full feature deck",
  description: "An editable 11-slide tour of text, shapes, groups, images, tables, charts, and backgrounds.",
  source: fullDeck.source.content,
  create: (source = fullDeck.source.content) => createExamplePresentation(parseExampleSource(source)),
};

const fixtureTitles: Record<string, string> = {
  text: "Rich text & notes",
  "theme-layout": "Themes & layouts",
  geometry: "Shapes & connectors",
  images: "Images & cropping",
  table: "Merged table cells",
  charts: "Bar, line & pie charts",
  "missing-asset": "Missing asset diagnostics",
};

function advancedCharts() {
  const deck = createPresentation({ metadata: { title: "Advanced charts" } });
  const box = { x: 90, y: 55, width: 780, height: 430 };
  const categories = ["Jan", "Feb", "Mar", "Apr"];
  const series = [{ name: "Orders", values: [24, 32, 18, 38] }];
  const points = { name: "Stores", xValues: [1, 3, 4, 7], yValues: [2, 5, 4, 8] };
  deck.addSlide().addElement({ type: "chart", chartType: "doughnut", title: "Orders by month", box, categories, series });
  deck.addSlide().addElement({ type: "chart", chartType: "area", title: "Order trend", box, categories, series });
  deck.addSlide().addElement({ type: "chart", chartType: "scatter", title: "Visits and sales", box, series: [points] });
  deck.addSlide().addElement({ type: "chart", chartType: "bubble", title: "Visits, sales, and store size", box, series: [{ ...points, sizes: [3, 16, 8, 12] }] });
  deck.addSlide().addElement({
    type: "chart", chartType: "combo", title: "Orders and revenue", box, categories,
    series: [{ ...series[0]!, kind: "bar" }, { kind: "line", name: "Revenue", values: [450, 620, 320, 740] }],
    axes: { value: { scale: { min: 0, max: 40, majorUnit: 10 } }, secondaryValue: { scale: { min: 0, max: 800, majorUnit: 200 } } },
  });
  return deck;
}

function transitions() {
  const deck = createPresentation({ metadata: { title: "Fade transition" } });
  const box = { x: 100, y: 200, width: 760, height: 100 };
  deck.addSlide().addElement({ type: "text", content: "Start the slide show", box });
  deck.addSlide({ transition: { type: "fade", durationMs: 800, onClick: true } })
    .addElement({ type: "text", content: "This slide fades in", box });
  return deck;
}

export const examples = [
  ...fixtureCatalog.map(({ id, capabilities }) => ({
    id, title: fixtureTitles[id] ?? id,
    description: capabilities.map((value) => value.replaceAll("-", " ")).join(" · "),
    create: () => createFixture(id, createPresentation),
  })),
  { id: "advanced-charts", title: "Advanced charts", description: "Doughnut, area, scatter, bubble, and dual-axis combination charts.", create: advancedCharts },
  { id: "transitions", title: "Fade transition", description: "Open the downloaded deck in PowerPoint and start a slide show to see the fade.", create: transitions },
];

export async function loadNativeSample() {
  const response = await fetch(nativeSampleUrl);
  if (!response.ok) throw new Error(`Could not load the sample (${response.status}).`);
  return new Uint8Array(await response.arrayBuffer());
}
