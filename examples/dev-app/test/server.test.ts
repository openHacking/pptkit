import { describe, expect, it } from "vitest";
import { createWorkbenchRequestHandler } from "../src/server";

async function invoke(pathname: string) {
  const handler = createWorkbenchRequestHandler();
  let statusCode = 200;
  let headers: Record<string, string> = {};
  let body = "";

  await handler(
    {
      url: pathname,
      headers: {
        host: "localhost:3210",
      },
    },
    {
      writeHead(nextStatusCode: number, nextHeaders: Record<string, string>) {
        statusCode = nextStatusCode;
        headers = nextHeaders;
        return this;
      },
      end(chunk?: string | Buffer) {
        body = String(chunk ?? "");
        return this;
      },
    } as never,
  );

  return {
    statusCode,
    headers,
    body,
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
    };

    expect(exampleResponse.statusCode).toBe(200);
    expect(examplePayload.example.id).toBe(workbenchPayload.examples[0]!.id);
    expect(examplePayload.diagnostics.length).toBeGreaterThan(0);

    const healthResponse = await invoke("/api/health");
    const healthPayload = JSON.parse(healthResponse.body) as { ok: boolean };

    expect(healthResponse.statusCode).toBe(200);
    expect(healthPayload.ok).toBe(true);
  });
});
