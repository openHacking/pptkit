import type { ExampleDefinition, ExampleSummary, FeatureId } from "./example-types.js";
import { saasHuntSwissStyleExample } from "./examples/saas-hunt-swiss-style.js";

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
    id: "export-pitch-deck",
    title: "Export Pitch Deck",
    feature: "export-pptx",
    description: "Exercise the real PPTX export path with a scenario-rich multi-slide case.",
    inputKind: "presentation-config",
    source: {
      label: "Pitch deck outline",
      content: `{
  "title": "Pitch Deck",
  "slides": [
    { "title": "Problem", "elements": ["Manual deck work is slow"] },
    { "title": "Solution", "elements": ["PPTKit automates structured output"] },
    { "title": "Roadmap", "elements": ["Preview", "Export", "Editing"] }
  ]
}`,
    },
    scenarioTags: ["pitch-deck", "founder-story"],
    expectedCapabilities: {
      normalize: "implemented",
      render: "placeholder",
      exportPptx: "implemented",
    },
    status: "ready",
    createInput() {
      return {
        title: "Pitch Deck",
        summary: "A higher-context scenario keeps export work grounded in a realistic multi-slide story.",
        slides: [
          {
            title: "Problem",
            elements: [
              "Manual deck work is slow",
              "Reformatting steals time from content",
            ],
          },
          {
            title: "Solution",
            elements: [
              "PPTKit automates structured output",
              "One IR supports preview and export",
            ],
          },
          {
            title: "Roadmap",
            elements: [
              "Milestone: Preview workbench",
              "Milestone: PPTX export",
              "Milestone: Editing workflows",
            ],
          },
        ],
      };
    },
  },
  {
    id: "export-mixed-media-slide",
    title: "Export Mixed Media Slide",
    feature: "export-pptx",
    description: "Prove the dev workbench can author and export text, shapes, and images from one source payload.",
    inputKind: "presentation-config",
    source: {
      label: "Mixed media slide config",
      content: `{
  "title": "Mixed Media Demo",
  "summary": "A single slide combining text, shapes, and an image asset.",
  "slides": [
    {
      "title": "Overview",
      "elements": [
        {
          "type": "text",
          "text": "PPTKit visual elements",
          "box": { "x": 48, "y": 40, "width": 360, "height": 28 },
          "style": { "fontSize": 24, "fontWeight": "bold" }
        },
        {
          "type": "shape",
          "shape": "rect",
          "box": { "x": 48, "y": 92, "width": 240, "height": 120 },
          "style": { "fill": "#DCEBFF", "stroke": "#2563EB", "strokeWidth": 2 }
        },
        {
          "type": "image",
          "altText": "Blue accent pixel",
          "asset": {
            "id": "accent-pixel",
            "source": {
              "type": "url",
              "value": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WnR8WQAAAAASUVORK5CYII="
            },
            "mimeType": "image/png",
            "altText": "Blue accent pixel"
          },
          "box": { "x": 320, "y": 92, "width": 180, "height": 180 }
        },
        {
          "type": "shape",
          "shape": "line",
          "box": { "x": 48, "y": 236, "width": 452, "height": 4 },
          "style": { "stroke": "#0F172A", "strokeWidth": 2 }
        },
        "Caption: Mixed element authoring now reaches the exporter."
      ]
    }
  ]
}`,
    },
    scenarioTags: ["mixed-media", "export"],
    expectedCapabilities: {
      normalize: "implemented",
      render: "placeholder",
      exportPptx: "implemented",
    },
    status: "ready",
    createInput() {
      return {
        title: "Mixed Media Demo",
        summary: "A single slide combining text, shapes, and an image asset.",
        slides: [
          {
            title: "Overview",
            elements: [
              {
                type: "text",
                text: "PPTKit visual elements",
                box: { x: 48, y: 40, width: 360, height: 28 },
                style: { fontSize: 24, fontWeight: "bold" },
              },
              {
                type: "shape",
                shape: "rect",
                box: { x: 48, y: 92, width: 240, height: 120 },
                style: { fill: "#DCEBFF", stroke: "#2563EB", strokeWidth: 2 },
              },
              {
                type: "image",
                altText: "Blue accent pixel",
                asset: {
                  id: "accent-pixel",
                  source: {
                    type: "url",
                    value:
                      "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WnR8WQAAAAASUVORK5CYII=",
                  },
                  mimeType: "image/png",
                  altText: "Blue accent pixel",
                },
                box: { x: 320, y: 92, width: 180, height: 180 },
              },
              {
                type: "shape",
                shape: "line",
                box: { x: 48, y: 236, width: 452, height: 4 },
                style: { stroke: "#0F172A", strokeWidth: 2 },
              },
              "Caption: Mixed element authoring now reaches the exporter.",
            ],
          },
        ],
      };
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
