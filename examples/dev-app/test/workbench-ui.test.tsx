// @vitest-environment jsdom

import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
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
        exportPptx: "implemented",
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
      status: "written",
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
      vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
        const url = String(input);

        if (url.endsWith("/report") && init?.method === "POST") {
          const body = JSON.parse(String(init.body)) as { source: string };
          if (body.source === "not json") {
            return new Response(JSON.stringify({ error: "Not found" }), { status: 404 });
          }
          const nextReport = buildReport("export-alpha", "export", "Applied Source");
          nextReport.example.source.content = body.source;
          nextReport.visualPreview.slides[0]!.title = "Applied Source";
          return new Response(JSON.stringify(nextReport), { status: 200 });
        }

        if (url.endsWith("/export") && init?.method === "POST") {
          return new Response("pptx-bytes", {
            status: 200,
            headers: { "content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
          });
        }

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

  it("applies edited source before exporting it", async () => {
    const user = userEvent.setup();
    const createObjectURL = vi.fn(() => "blob:export");
    const revokeObjectURL = vi.fn();
    vi.stubGlobal("URL", { createObjectURL, revokeObjectURL });
    render(<App />);

    expect(await screen.findByText("PPTKit Workbench")).toBeTruthy();
    const source = screen.getByRole("textbox", { name: "Import Alpha.md" });
    await user.clear(source);
    fireEvent.change(source, {
      target: { value: '{"title":"Updated","slides":[{"title":"New slide","elements":["New line"]}]}' },
    });

    expect(screen.getByText("Unsaved changes")).toBeTruthy();
    expect((screen.getByRole("button", { name: "Export in browser" }) as HTMLButtonElement).disabled).toBe(true);
    expect((screen.getByRole("button", { name: "Export via server" }) as HTMLButtonElement).disabled).toBe(true);

    await user.click(screen.getByRole("button", { name: "Apply changes" }));
    await waitFor(() => {
      expect(screen.getAllByText("Applied Source").length).toBeGreaterThan(0);
    });

    await user.click(screen.getByRole("button", { name: "Export in browser" }));
    await waitFor(() => {
      expect(createObjectURL).toHaveBeenCalledOnce();
      expect(screen.getByText("PPTX exported in browser successfully.")).toBeTruthy();
    });
    expect(vi.mocked(fetch).mock.calls.some(([url]) => String(url).endsWith("/export"))).toBe(false);

    await user.click(screen.getByRole("button", { name: "Export via server" }));
    await waitFor(() => {
      expect(createObjectURL).toHaveBeenCalledTimes(2);
      expect(screen.getByText("PPTX exported via server successfully.")).toBeTruthy();
    });
    expect(vi.mocked(fetch).mock.calls.filter(([url]) => String(url).endsWith("/export"))).toHaveLength(1);
  });

  it("keeps the applied report when source JSON is invalid", async () => {
    const user = userEvent.setup();
    render(<App />);

    expect(await screen.findByText("PPTKit Workbench")).toBeTruthy();
    const source = screen.getByRole("textbox", { name: "Import Alpha.md" });
    await user.clear(source);
    await user.type(source, "not json");
    await user.click(screen.getByRole("button", { name: "Apply changes" }));

    await waitFor(() => {
      expect(screen.getByText("Not found")).toBeTruthy();
    });
    expect(screen.getByText("Import Alpha description")).toBeTruthy();
  });
});
