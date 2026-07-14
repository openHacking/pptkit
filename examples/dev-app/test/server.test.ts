import { describe, expect, it } from "vitest";
import { createWorkbenchRequestHandler } from "../src/server";

async function invoke(pathname: string, method = "GET", body?: string) {
  const handler = createWorkbenchRequestHandler();
  let statusCode = 200;
  let headers: Record<string, string> = {};
  let responseBody = "";

  const request = {
      url: pathname,
      method,
      headers: {
        host: "localhost:3210",
      },
      async *[Symbol.asyncIterator]() {
        if (body !== undefined) {
          yield body;
        }
      },
    };

  await handler(
    request,
    {
      writeHead(nextStatusCode: number, nextHeaders: Record<string, string>) {
        statusCode = nextStatusCode;
        headers = nextHeaders;
        return this;
      },
      end(chunk?: string | Buffer) {
        responseBody = String(chunk ?? "");
        return this;
      },
    } as never,
  );

  return {
    statusCode,
    headers,
    body: responseBody,
  };
}

describe("workbench server", () => {
  it("serves workbench payload and example reports", async () => {
    const workbenchResponse = await invoke("/api/workbench");
    const workbenchPayload = JSON.parse(workbenchResponse.body) as {
      features: string[];
      examples: Array<{ id: string }>;
    };

    expect(workbenchResponse.statusCode).toBe(200);
    expect(workbenchPayload.features.length).toBeGreaterThan(0);
    expect(workbenchPayload.examples.length).toBeGreaterThan(0);

    const exampleResponse = await invoke(`/api/examples/${workbenchPayload.examples[0]!.id}`);
    const examplePayload = JSON.parse(exampleResponse.body) as {
      example: { id: string };
      diagnostics: string[];
      exportResult: { output: string; status: string };
    };

    expect(exampleResponse.statusCode).toBe(200);
    expect(examplePayload.example.id).toBe(workbenchPayload.examples[0]!.id);
    expect(examplePayload.diagnostics.length).toBeGreaterThan(0);
    expect(examplePayload.exportResult.status).toBe("not-exported");
    expect(examplePayload.exportResult.output).toBe("");

    const healthResponse = await invoke("/api/health");
    const healthPayload = JSON.parse(healthResponse.body) as { ok: boolean };

    expect(healthResponse.statusCode).toBe(200);
    expect(healthPayload.ok).toBe(true);
  });

  it("rebuilds reports and exports applied source", async () => {
    const source = JSON.stringify({
      title: "Applied deck",
      slides: [{ title: "Changed", elements: ["A", "B"] }],
    });
    const reportResponse = await invoke("/api/examples/export-full-feature-deck/report", "POST", JSON.stringify({ source }));
    const reportPayload = JSON.parse(reportResponse.body) as {
      normalizedDocument: { title: string; slideCount: number };
    };

    expect(reportResponse.statusCode).toBe(200);
    expect(reportPayload.normalizedDocument.title).toBe("Applied deck");
    expect(reportPayload.normalizedDocument.slideCount).toBe(1);

    const exportResponse = await invoke("/api/examples/export-full-feature-deck/export", "POST", JSON.stringify({ source }));
    expect(exportResponse.statusCode).toBe(200);
    expect(exportResponse.headers["content-type"]).toContain("presentationml.presentation");
    expect(exportResponse.headers["content-disposition"]).toContain("export-full-feature-deck.pptx");
    expect(exportResponse.body.length).toBeGreaterThan(0);
  });

  it("exports the default full-feature deck", async () => {
    const exampleResponse = await invoke("/api/examples/export-full-feature-deck");
    const examplePayload = JSON.parse(exampleResponse.body) as {
      example: { source: { content: string } };
      normalizedDocument: { slideCount: number };
    };

    expect(exampleResponse.statusCode).toBe(200);
    expect(examplePayload.normalizedDocument.slideCount).toBe(8);

    const exportResponse = await invoke(
      "/api/examples/export-full-feature-deck/export",
      "POST",
      JSON.stringify({ source: examplePayload.example.source.content }),
    );

    expect(exportResponse.statusCode).toBe(200);
    expect(exportResponse.body.length).toBeGreaterThan(0);
  });

  it("rejects invalid source requests", async () => {
    const response = await invoke(
      "/api/examples/export-full-feature-deck/report",
      "POST",
      JSON.stringify({ source: "not json" }),
    );

    expect(response.statusCode).toBe(400);
    expect(JSON.parse(response.body).error).toContain("valid JSON");
  });
});
