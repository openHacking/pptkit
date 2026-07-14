import type { ExampleDefinition, ExampleInputData, ExampleSummary, FeatureId } from "./example-types.js";
import { saasHuntSwissStyleExample } from "./examples/saas-hunt-swiss-style.js";

const exportFullFeatureDeckInput: ExampleInputData = {
  title: "PPTKit Export Full Feature Deck",
  summary: "Each slide isolates one or two exporter features and labels them in the slide content for quick export verification.",
  slides: [
    {
      id: "text-structure",
      title: "Feature: slide structure + plain text",
      elements: [
        "Feature: slide title and ordered multi-slide structure",
        "Feature: plain text element — verify readable content and line spacing",
      ],
    },
    {
      id: "styled-text",
      title: "Feature: styled text",
      elements: [
        {
          type: "text",
          text: "Feature: bold, color, size, alignment, and auto-fit",
          box: { x: 48, y: 72, width: 624, height: 48 },
          style: {
            fontSize: 24,
            fontWeight: "bold",
            color: "#2563EB",
            align: "center",
            autoFit: { mode: "shrink", fontScale: 0.8 },
          },
        },
        "Export check: styled text should remain legible and centered.",
      ],
    },
    {
      id: "basic-shapes",
      title: "Feature: rectangle + ellipse shapes",
      elements: [
        {
          type: "shape",
          shape: "rect",
          box: { x: 48, y: 72, width: 260, height: 160 },
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
      render: "placeholder",
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
      render: "placeholder",
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
      render: "placeholder",
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
      render: "placeholder",
      exportPptx: "implemented",
    },
    status: "ready",
    createInput() {
      return exportFullFeatureDeckInput;
    },
  },
  saasHuntSwissStyleExample,
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
