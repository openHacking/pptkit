import { createPresentation, validatePresentation } from "@pptkit/core";
import { writePptx } from "@pptkit/pptx-exporter/node";

const presentation = createPresentation({
  metadata: { title: "Hello PPTKit", author: "Example Team" },
  theme: { colors: { accent1: "2457D6" } },
});

const slide = presentation.addSlide();
slide.addElement({
  type: "text",
  content: [{
    runs: [
      { text: "Hello ", style: { fontSize: 36 } },
      {
        text: "PPTKit",
        style: { fontSize: 36, bold: true, color: { theme: "accent1" } },
      },
    ],
  }],
  box: { x: 64, y: 64, width: 520, height: 72 },
});

const diagnostics = validatePresentation(presentation);
if (diagnostics.some((diagnostic) => diagnostic.severity === "error")) {
  throw new Error(JSON.stringify(diagnostics, null, 2));
}

const result = await writePptx(presentation, {
  output: "./hello-pptkit.pptx",
});

console.log(result.status, result.output, result.warnings);
