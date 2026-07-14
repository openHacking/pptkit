import type { ExampleDefinition, ExampleInputData, ExampleSummary, FeatureId } from "./example-types.js";
import { exportgptProExample } from "./examples/exportgpt-pro.js";
import { saasHuntSwissStyleExample } from "./examples/saas-hunt-swiss-style.js";

const exportFullFeatureDeckInput: ExampleInputData = {
  title: "PPTKit Export Full Feature Deck",
  summary: "Each slide isolates one or two exporter features and labels them in the slide content for quick export verification.",
  textStylePresets: {
    featureText: {
      frame: { margin: 0, verticalAlign: "middle" },
      paragraph: { align: "center" },
      run: { fontSize: 24, bold: true, color: "1D4ED8" },
    },
    shapeLabel: {
      frame: { margin: 0, verticalAlign: "middle" },
      paragraph: { align: "center" },
      run: { fontSize: 20, bold: true, color: "0F172A" },
    },
  },
  slides: [
    {
      id: "text-structure",
      title: "Feature: plain text + bullet list",
      elements: [
        "Feature: plain text element — verify readable content and line spacing",
        {
          type: "text",
          text: "Feature: bullet paragraph — verify bullet character and spacing",
          box: { x: 48, y: 112, width: 624, height: 32 },
          style: {
            fontSize: 20,
            bullet: { type: "bullet", character: "•" },
            spaceAfter: 8,
          },
        },
      ],
    },
    {
      id: "styled-text",
      title: "Feature: character styles + numbered list",
      elements: [
        {
          type: "text",
          text: "Feature: bold, italic, underline, strike, color, size, and auto-fit",
          textStylePreset: "featureText",
          box: { x: 48, y: 72, width: 624, height: 48 },
          style: {
            fontSize: 24,
            fontFamily: "Arial",
            fontWeight: "bold",
            italic: true,
            underline: true,
            strike: true,
            color: "#2563EB",
            language: "en-US",
            align: "center",
            lineSpacing: 1.1,
            autoFit: { mode: "shrink", fontScale: 0.8 },
          },
        },
        {
          type: "text",
          text: "Feature: numbered paragraph — verify alpha numbering\nFirst numbered item\nSecond numbered item",
          box: { x: 48, y: 152, width: 624 },
          style: {
            fontSize: 20,
            bullet: { type: "number", style: "alphaLowerPeriod", startAt: 1 },
            indent: 18,
            hanging: 9,
          },
        },
      ],
    },
    {
      id: "basic-shapes",
      title: "Feature: native shape text + text style preset",
      elements: [
        {
          type: "shape",
          shape: "rect",
          box: { x: 48, y: 72, width: 260, height: 160 },
          text: {
            content: "Feature: editable text inside one rectangle shape",
            textStylePreset: "shapeLabel",
          },
          style: { fill: "#DCEBFF", stroke: "#2563EB", strokeWidth: 2 },
        },
        {
          type: "shape",
          shape: "ellipse",
          box: { x: 380, y: 92, width: 180, height: 120 },
          style: { fill: "#DCFCE7", stroke: "#16A34A", strokeWidth: 2 },
        },
      ],
    },
    {
      id: "group",
      title: "Feature: nested group",
      elements: [
        "Feature: group — verify local child coordinates and grouped export",
        {
          type: "group",
          box: { x: 160, y: 96, width: 400, height: 220 },
          coordinateSize: { width: 400, height: 220 },
          children: [
            {
              type: "shape",
              shape: "rect",
              box: { x: 24, y: 28, width: 160, height: 120 },
              style: { fill: "#DBEAFE", stroke: "#2563EB", strokeWidth: 2 },
            },
            {
              type: "shape",
              shape: "ellipse",
              box: { x: 216, y: 64, width: 140, height: 96 },
              style: { fill: "#DCFCE7", stroke: "#16A34A", strokeWidth: 2 },
            },
          ],
        },
      ],
    },
    {
      id: "table",
      title: "Feature: native editable table",
      elements: [
        "Feature: native table — verify editable cells, column widths, and colSpan",
        {
          type: "table",
          box: { x: 48, y: 104, width: 624, height: 220 },
          columns: [220, 180, 180],
          rows: [
            { cells: [{ content: "Quarterly results", colSpan: 3 }] },
            { cells: [{ content: "Workstream" }, { content: "Owner" }, { content: "Status" }] },
            { cells: [{ content: "Export" }, { content: "PPTKit" }, { content: "Ready" }] },
          ],
        },
      ],
    },
    {
      id: "connector",
      title: "Feature: connector line",
      elements: [
        "Feature: connector line — verify start/end coordinates and stroke",
        {
          type: "shape",
          shape: "line",
          box: { x: 48, y: 128, width: 576, height: 0 },
          style: { stroke: "#0F172A", strokeWidth: 3 },
        },
      ],
    },
    {
      id: "image-asset",
      title: "Feature: image asset + alt text",
      elements: [
        "Feature: image asset registered from a data URL with accessibility text",
        {
          type: "image",
          altText: "Blue accent pixel used to verify image export",
          asset: {
            id: "accent-pixel",
            source: {
              type: "url",
              value: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
            },
            mimeType: "image/png",
            altText: "Blue accent pixel used to verify image export",
          },
          box: { x: 48, y: 104, width: 220, height: 180 },
        },
      ],
    },
    {
      id: "background-and-summary",
      title: "Feature: slide background + mixed content",
      background: "#FFF7ED",
      elements: [
        {
          type: "text",
          text: "Feature: solid slide background",
          box: { x: 48, y: 72, width: 624, height: 32 },
          style: { fontSize: 22, fontWeight: "bold", color: "#9A3412" },
        },
        "Feature: final multi-slide export summary — review all previous labels in one PPTX.",
      ],
    },
  ],
};

const exampleDefinitions: ExampleDefinition[] = [
  {
    id: "import-html-hero",
    title: "HTML Hero Section",
    feature: "import-html",
    description: "Validate how structured HTML authoring input is turned into a presentation-shaped document.",
    inputKind: "html",
    source: {
      label: "Landing page fragment",
      content: `<section>
  <h1>PPTKit Quarterly Launch</h1>
  <p>Convert polished HTML blocks into draft slides.</p>
  <ul>
    <li>Hero headline and supporting copy</li>
    <li>Three launch metrics</li>
    <li>CTA-ready closing slide</li>
  </ul>
</section>`,
    },
    scenarioTags: ["basic-deck", "marketing"],
    expectedCapabilities: {
      normalize: "implemented",
      render: "implemented",
      exportPptx: "implemented",
    },
    status: "ready",
    createInput() {
      return {
        title: "Quarterly Launch",
        summary: "HTML blocks are normalized into a concise presentation draft for marketing review.",
        slides: [
          {
            title: "PPTKit Quarterly Launch",
            elements: [
              "Headline: PPTKit Quarterly Launch",
              "Body: Convert polished HTML blocks into draft slides.",
              "List: Hero headline and supporting copy / Three launch metrics / CTA-ready closing slide",
            ],
          },
          {
            title: "Launch CTA",
            elements: [
              "Action: Review generated deck structure",
              "Next step: Compare preview with future HTML importer",
            ],
          },
        ],
      };
    },
  },
  {
    id: "import-markdown-status",
    title: "Markdown Status Update",
    feature: "import-markdown",
    description: "Check that markdown sections map cleanly to slide-level structure and narrative order.",
    inputKind: "markdown",
    source: {
      label: "Team update markdown",
      content: `# Platform Weekly

## Wins
- Preview workbench drafted
- Smoke tests passing

## Risks
- Renderer package not started
- Rich PPTX feature coverage is still limited

## Ask
- Align on functional test taxonomy`,
    },
    scenarioTags: ["basic-deck", "status-update"],
    expectedCapabilities: {
      normalize: "implemented",
      render: "implemented",
      exportPptx: "implemented",
    },
    status: "ready",
    createInput() {
      return {
        title: "Platform Weekly",
        summary: "Markdown headings and bullets are promoted into slide candidates for team communication.",
        slides: [
          {
            title: "Platform Weekly",
            elements: [
              "Heading: Platform Weekly",
              "Theme: Weekly engineering status",
            ],
          },
          {
            title: "Wins",
            elements: [
              "Bullet: Preview workbench drafted",
              "Bullet: Smoke tests passing",
            ],
          },
          {
            title: "Risks and Ask",
            elements: [
              "Risk: Renderer package not started",
              "Risk: Rich PPTX feature coverage is still limited",
              "Ask: Align on functional test taxonomy",
            ],
          },
        ],
      };
    },
  },
  {
    id: "render-product-brief",
    title: "Render Product Brief",
    feature: "render",
    description: "Exercise the normalized document path and structural preview without depending on PPTX output.",
    inputKind: "presentation-config",
    source: {
      label: "JSON-like authoring spec",
      content: `{
  "title": "Product Brief",
  "slides": [
    { "title": "Goal", "elements": ["Clarify preview architecture"] },
    { "title": "Flow", "elements": ["IR -> render", "IR -> export"] }
  ]
}`,
    },
    scenarioTags: ["basic-deck", "product-brief"],
    expectedCapabilities: {
      normalize: "implemented",
      render: "implemented",
      exportPptx: "implemented",
    },
    status: "ready",
    createInput() {
      return {
        title: "Product Brief",
        summary: "A direct presentation spec keeps the workbench useful before dedicated importers exist.",
        slides: [
          {
            title: "Goal",
            elements: [
              "Clarify preview architecture",
              "Keep authored state separate from derived output",
            ],
          },
          {
            title: "Flow",
            elements: [
              "Path: IR -> structural preview",
              "Path: IR -> PPTX export",
            ],
          },
        ],
      };
    },
  },
  {
    id: "export-full-feature-deck",
    title: "Export Full Feature Deck",
    feature: "export-pptx",
    description: "Exercise the complete PPTX export path with one labeled slide per feature pair.",
    inputKind: "presentation-config",
    source: {
      label: "Full feature deck config",
      content: JSON.stringify(exportFullFeatureDeckInput, null, 2),
    },
    scenarioTags: ["full-coverage", "pitch-deck", "mixed-media", "export"],
    expectedCapabilities: {
      normalize: "implemented",
      render: "implemented",
      exportPptx: "implemented",
    },
    status: "ready",
    createInput() {
      return exportFullFeatureDeckInput;
    },
  },
  saasHuntSwissStyleExample,
  exportgptProExample,
];

export function listExamples(): ExampleDefinition[] {
  return exampleDefinitions;
}

export function getExampleById(id: string): ExampleDefinition | undefined {
  return exampleDefinitions.find((example) => example.id === id);
}

export function listExampleSummaries(): ExampleSummary[] {
  return exampleDefinitions.map((example) => ({
    id: example.id,
    title: example.title,
    feature: example.feature,
    description: example.description,
    inputKind: example.inputKind,
    scenarioTags: [...example.scenarioTags],
    status: example.status,
  }));
}

export function listFeatures(): FeatureId[] {
  return [...new Set(exampleDefinitions.map((example) => example.feature))];
}
