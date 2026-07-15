import { createPresentation } from "@pptkit/core";
import { generatePptx } from "@pptkit/pptx-exporter";
import { renderPresentationToSvg } from "@pptkit/svg-renderer";

const presentation = createPresentation({
  metadata: { title: "PPTKit Browser ESM", author: "PPTKit" },
  theme: { colors: { accent1: "2457D6", accent2: "11A683" } },
});

presentation.addSlide({
  elements: [
    {
      type: "text",
      content: "PPTKit runs in your browser",
      box: { x: 64, y: 72, width: 760, height: 72 },
      style: { run: { fontSize: 38, bold: true, color: { theme: "text1" } } },
    },
    {
      type: "shape",
      shape: "roundRect",
      box: { x: 64, y: 190, width: 832, height: 210 },
      style: { fill: { type: "solid", color: { theme: "accent1" }, opacity: 0.12 } },
      text: {
        content: [{ runs: [{ text: "No Node.js runtime. No server-side PPTX generation.", style: { fontSize: 26 } }] }],
        frame: { margin: 28, verticalAlign: "middle" },
      },
    },
  ],
});

presentation.addSlide({
  background: { type: "solid", color: "10213F" },
  elements: [
    {
      type: "text",
      content: "One document, two browser outputs",
      box: { x: 64, y: 72, width: 800, height: 64 },
      style: { run: { fontSize: 34, bold: true, color: "FFFFFF" } },
    },
    {
      type: "text",
      content: "SVG preview → editable PPTX download",
      box: { x: 64, y: 190, width: 800, height: 54 },
      style: { run: { fontSize: 26, color: { theme: "accent2" } } },
    },
  ],
});

const previewButton = document.querySelector("#preview-button");
const downloadButton = document.querySelector("#download-button");
const preview = document.querySelector("#preview");
const status = document.querySelector("#status");
const warnings = document.querySelector("#warnings");

function showWarnings(items) {
  warnings.textContent = items.length === 0
    ? "No warnings."
    : items.map((warning) => `${warning.code}: ${warning.message}`).join("\n");
}

async function run(button, task) {
  button.disabled = true;
  warnings.textContent = "";
  try {
    await task();
  } catch (error) {
    status.textContent = `Failed: ${error instanceof Error ? error.message : String(error)}`;
  } finally {
    button.disabled = false;
  }
}

previewButton.addEventListener("click", () => run(previewButton, async () => {
  status.textContent = "Rendering SVG previews…";
  const result = await renderPresentationToSvg(presentation);
  preview.replaceChildren(...result.slides.map((slide) => {
    const container = document.createElement("article");
    container.className = "slide";
    container.innerHTML = slide.svg;
    return container;
  }));
  showWarnings(result.warnings);
  status.textContent = `Rendered ${result.slides.length} slides in the browser.`;
}));

downloadButton.addEventListener("click", () => run(downloadButton, async () => {
  status.textContent = "Generating PPTX…";
  const result = await generatePptx(presentation);
  const blob = new Blob([result.bytes.slice().buffer], {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "pptkit-browser-example.pptx";
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 0);
  showWarnings(result.warnings);
  status.textContent = `Generated ${result.slideCount} slides (${result.byteLength} bytes).`;
}));
