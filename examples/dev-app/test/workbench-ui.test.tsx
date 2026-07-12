// @vitest-environment jsdom

import { cleanup, render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { ExampleReport, WorkbenchPayload } from "../src/example-types";
import App from "../src/client/App";

const payload: WorkbenchPayload = {
  features: ["import", "export"],
  examples: [
    {
      id: "import-alpha",
      title: "Import Alpha",
      feature: "import",
      description: "Import example alpha",
      inputKind: "markdown",
      scenarioTags: ["fixtures"],
      status: "ready",
    },
    {
      id: "import-beta",
      title: "Import Beta",
      feature: "import",
      description: "Import example beta",
      inputKind: "html",
      scenarioTags: ["preview"],
      status: "placeholder",
    },
    {
      id: "export-alpha",
      title: "Export Alpha",
      feature: "export",
      description: "Export example alpha",
      inputKind: "presentation-config",
      scenarioTags: ["export"],
      status: "ready",
    },
  ],
};

function buildReport(id: string, feature: string, title: string): ExampleReport {
  return {
    example: {
      id,
      title,
      feature,
      description: `${title} description`,
      inputKind: "markdown",
      source: {
        label: `${title}.md`,
        content: `# ${title}`,
      },
      scenarioTags: ["fixtures", feature],
      expectedCapabilities: {
        normalize: "implemented",
        render: "placeholder",
        exportPptx: "not-implemented",
      },
      status: "ready",
      createInput() {
        return {
          title,
          summary: `${title} summary`,
          slides: [
            {
              title,
              elements: ["One", "Two"],
            },
          ],
        };
      },
    },
    normalizedDocument: {
      title,
      summary: `${title} summary`,
      slideCount: 1,
      slides: [{ id: `${id}-1`, title, elements: ["One", "Two"] }],
    },
    visualPreview: {
      status: "structural-preview",
      slides: [{ id: `${id}-slide`, title, body: ["One", "Two"] }],
    },
    renderResult: {
      slideCount: 1,
      status: "placeholder",
    },
    exportResult: {
      output: `${id}.pptx`,
      slideCount: 1,
      status: "not-implemented",
    },
    diagnostics: [`${title} diagnostic`],
  };
}

const reports: Record<string, ExampleReport> = {
  "import-alpha": buildReport("import-alpha", "import", "Import Alpha"),
  "import-beta": buildReport("import-beta", "import", "Import Beta"),
  "export-alpha": buildReport("export-alpha", "export", "Export Alpha"),
};

describe("workbench app", () => {
  beforeEach(() => {
    vi.stubGlobal(
      "ResizeObserver",
      class ResizeObserver {
        observe() {}
        unobserve() {}
        disconnect() {}
      },
    );

    vi.stubGlobal(
      "fetch",
      vi.fn(async (input: RequestInfo | URL) => {
        const url = String(input);

        if (url.endsWith("/api/workbench")) {
          return new Response(JSON.stringify(payload), {
            status: 200,
            headers: { "content-type": "application/json" },
          });
        }

        const exampleId = url.split("/api/examples/")[1];

        if (exampleId !== undefined && reports[exampleId] !== undefined) {
          return new Response(JSON.stringify(reports[exampleId]), {
            status: 200,
            headers: { "content-type": "application/json" },
          });
        }

        return new Response(JSON.stringify({ error: "Not found" }), {
          status: 404,
          headers: { "content-type": "application/json" },
        });
      }),
    );
  });

  afterEach(() => {
    cleanup();
    vi.unstubAllGlobals();
  });

  it("loads the default example and updates when feature and example selections change", async () => {
    const user = userEvent.setup();
    render(<App />);

    expect(await screen.findByText("PPTKit Workbench")).toBeTruthy();
    expect(screen.getByText("Import Alpha description")).toBeTruthy();
    expect(screen.queryByText(/^Input:/)).toBeNull();
    expect(screen.queryByText("fixtures")).toBeNull();

    await user.click(screen.getByRole("button", { name: /Import Beta/ }));
    await waitFor(() => {
      expect(screen.getByText("Import Beta description")).toBeTruthy();
    });

    await user.click(screen.getByRole("tab", { name: "export" }));
    await waitFor(() => {
      expect(screen.getByText("Export Alpha description")).toBeTruthy();
    });

    expect(screen.getByText("Export Alpha diagnostic")).toBeTruthy();
  });
});
