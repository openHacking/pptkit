// Prepare independent combination authoring and source edit-save samples.
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { createZip } from "../packages/pptx-exporter/dist/archive/create-zip.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const edited = process.argv.includes("--edited");
const publicApi = edited || process.argv.includes("--public");
const [destination] = process.argv.slice(2).filter(arg => arg !== "--public" && arg !== "--edited");
const out = path.resolve(destination ?? (publicApi ? ".tmp/combination-public-probe" : ".tmp/combination-chart-probe"));
await mkdir(out, { recursive: true });
const categories = ["Jan", "Feb", "Mar"];
const bar = { name: "Volume", color: "#4472C4", values: [2, 5, 8] };
const line = { name: "Rate", color: "#ED7D31", values: [800, 200, 600] };
const deck = createPresentation({ size: { width: 720, height: 405 } });
const common = { type: "chart", title: "Two-axis probe", box: { x: 45, y: 30, width: 600, height: 345 }, categories, legend: { visible: false } };
if (publicApi) {
  deck.addSlide({ elements: [{ ...common, chartType: "combo",
    series: [{ ...bar, kind: "bar" }, { ...line, kind: "line" }],
    axes: { value: { scale: { min: 0, max: 10, majorUnit: 2 } }, secondaryValue: { scale: { min: 0, max: 1000, majorUnit: 200 } } } }] });
  const { bytes } = await generatePptx(deck);
  await writeFile(path.join(out, "combination.pptx"), bytes);
  if (edited) {
    const { importPptx } = await import("../packages/pptx-parser/dist/index.js");
    const session = await importPptx(bytes, { projection: "content", chartEditSaveProfile: "existing-values-v3" });
    try {
      const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]);
      if (chart.chartType !== "combo" || chart.series[0].values[1] !== 5 || chart.series[1].values[1] !== 200) throw new Error("Unexpected combination baseline.");
      chart.series[0].values[1] = 7;
      chart.series[1].values[1] = 400;
      slide.replaceElement(chart.id, chart);
      const candidate = await session.prepareSave({ mode: "edit" });
      if (candidate.status !== "ready") throw new Error(JSON.stringify(candidate.report));
      await writeFile(path.join(out, "combination-edited.pptx"), (await session.save(candidate.plan)).bytes);
    } finally { session.dispose(); }
  }
} else {
  deck.addSlide({ elements: [{ ...common, chartType: "bar", orientation: "vertical", series: [bar],
    axes: { value: { scale: { min: 0, max: 10, majorUnit: 2 } } } }] });
  deck.addSlide({ elements: [{ ...common, chartType: "line", series: [line],
    axes: { value: { scale: { min: 0, max: 1000, majorUnit: 200 } } } }] });
  const parts = readGeneratedZip((await generatePptx(deck)).bytes);
  const barChart = parts.get("ppt/charts/chart1.xml").toString();
  const lineChart = parts.get("ppt/charts/chart2.xml").toString();
  const element = (xml, tag) => {
    const match = xml.match(new RegExp(`<c:${tag}>[\\s\\S]*?<\\/c:${tag}>`));
    if (!match) throw new Error(`Missing ${tag}`);
    return match[0];
  };
  const replaceRequired = (xml, search, replacement) => {
    const changed = xml.replace(search, replacement);
    if (changed === xml) throw new Error(`Unable to replace ${search}`);
    return changed;
  };
  const primary = element(barChart, "barChart"), cat = element(barChart, "catAx"), val = element(barChart, "valAx");
  let secondary = element(lineChart, "lineChart");
  secondary = replaceRequired(secondary, '<c:idx val="0"/><c:order val="0"/>', '<c:idx val="1"/><c:order val="1"/>');
  secondary = replaceRequired(secondary, /\$B\$/g, "$$C$$");
  secondary = replaceRequired(secondary, '<c:axId val="1"/><c:axId val="2"/>', '<c:axId val="3"/><c:axId val="4"/>');
  let cat2 = element(lineChart, "catAx");
  cat2 = replaceRequired(cat2, '<c:axId val="1"/>', '<c:axId val="3"/>');
  cat2 = replaceRequired(cat2, '<c:crossAx val="2"/>', '<c:crossAx val="4"/>');
  cat2 = replaceRequired(cat2, '<c:delete val="0"/>', '<c:delete val="1"/>');
  cat2 = replaceRequired(cat2, '<c:axPos val="b"/>', '<c:axPos val="t"/>');
  let val2 = element(lineChart, "valAx");
  val2 = replaceRequired(val2, '<c:axId val="2"/>', '<c:axId val="4"/>');
  val2 = replaceRequired(val2, '<c:crossAx val="1"/>', '<c:crossAx val="3"/>');
  val2 = replaceRequired(val2, '<c:axPos val="l"/>', '<c:axPos val="r"/>');
  val2 = replaceRequired(val2, '<c:crosses val="autoZero"/>', '<c:crosses val="max"/>');
  val2 = replaceRequired(val2, '<c:crossBetween val="midCat"/>', '<c:crossBetween val="between"/>');

  const carrier = createPresentation({ size: { width: 720, height: 405 } });
  carrier.addSlide({ elements: [{ ...common, chartType: "bar", series: [bar, line] }] });
  const result = readGeneratedZip((await generatePptx(carrier)).bytes);
  const target = "ppt/charts/chart1.xml";
  const combined = replaceRequired(barChart, primary + cat + val, primary + secondary + cat + val + cat2 + val2);
  result.set(target, new TextEncoder().encode(combined));
  await writeFile(path.join(out, "combination.pptx"), createZip([...result].map(([name, data]) => ({ name, data }))));
}
