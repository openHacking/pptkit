import { createServer, type IncomingMessage, type Server, type ServerResponse } from "node:http";
import { existsSync } from "node:fs";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { fileURLToPath } from "node:url";
import { getExampleById, listExamples, listExampleSummaries, listFeatures } from "./example-registry.js";
import { buildExampleReport } from "./report-builder.js";
import type { WorkbenchPayload } from "./example-types.js";
import { ExampleSourceError, parseExampleSource } from "./source-parser.js";
import { createPresentation } from "../../../packages/core/dist/index.js";
import { exportPptx } from "../../../packages/pptx-exporter/dist/index.js";

const sourceDir = path.dirname(fileURLToPath(import.meta.url));
const clientDir = path.resolve(sourceDir, "../client");
const port = Number(process.env.PORT ?? "3210");

function json(data: unknown) {
  return JSON.stringify(data, null, 2);
}

function contentType(filePath: string): string {
  if (filePath.endsWith(".css")) {
    return "text/css; charset=utf-8";
  }

  if (filePath.endsWith(".html")) {
    return "text/html; charset=utf-8";
  }

  if (filePath.endsWith(".js")) {
    return "text/javascript; charset=utf-8";
  }

  if (filePath.endsWith(".svg")) {
    return "image/svg+xml";
  }

  return "application/octet-stream";
}

async function serveFile(response: ServerResponse, filePath: string) {
  const content = await readFile(filePath);
  response.writeHead(200, { "content-type": contentType(filePath) });
  response.end(content);
}

async function serveApp(response: ServerResponse) {
  const entryFile = path.join(clientDir, "index.html");

  if (!existsSync(entryFile)) {
    response.writeHead(503, { "content-type": "application/json; charset=utf-8" });
    response.end(
      json({
        error: "Client build not found. Run `pnpm --filter @pptkit/dev-app build` or use `pnpm dev`.",
      }),
    );
    return;
  }

  await serveFile(response, entryFile);
}

async function readRequestBody(request: AsyncIterable<Buffer | string>): Promise<string> {
  const chunks: Buffer[] = [];

  for await (const chunk of request) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  }

  return Buffer.concat(chunks).toString("utf8");
}

function sourceFromBody(body: string): string {
  let value: unknown;

  try {
    value = JSON.parse(body);
  } catch {
    throw new ExampleSourceError("Request body must be valid JSON.");
  }

  if (typeof value !== "object" || value === null || typeof (value as { source?: unknown }).source !== "string") {
    throw new ExampleSourceError('Request body requires a string "source" field.');
  }

  return (value as { source: string }).source;
}

function buildWorkbenchPayload(): WorkbenchPayload {
  return {
    features: listFeatures(),
    examples: listExampleSummaries(),
  };
}

export function createWorkbenchServer() {
  return createServer(createWorkbenchRequestHandler());
}

export function createWorkbenchRequestHandler() {
  return async (
    request: Pick<IncomingMessage, "url" | "headers" | "method"> & AsyncIterable<Buffer | string>,
    response: ServerResponse,
  ) => {
    try {
      const hostHeader = Array.isArray(request.headers.host)
        ? request.headers.host[0]
        : request.headers.host;
      const url = new URL(request.url ?? "/", `http://${hostHeader ?? "localhost"}`);

      if (url.pathname === "/api/workbench") {
        response.writeHead(200, { "content-type": "application/json; charset=utf-8" });
        response.end(json(buildWorkbenchPayload()));
        return;
      }

      if (url.pathname.startsWith("/api/examples/")) {
        const pathParts = url.pathname.replace("/api/examples/", "").split("/");
        const id = pathParts[0] ?? "";
        const action = pathParts[1];
        const example = getExampleById(id);

        if (example === undefined) {
          response.writeHead(404, { "content-type": "application/json; charset=utf-8" });
          response.end(json({ error: `Unknown example: ${id}` }));
          return;
        }

        if (request.method === "POST" && (action === "report" || action === "export")) {
          const source = sourceFromBody(await readRequestBody(request));

          if (action === "report") {
            const report = await buildExampleReport(example, source);
            response.writeHead(200, { "content-type": "application/json; charset=utf-8" });
            response.end(json(report));
            return;
          }

          const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), "pptkit-export-"));
          const outputPath = path.join(temporaryDirectory, `${example.id}.pptx`);

          try {
            const input = parseExampleSource(source);
            const presentation = createPresentation({ title: input.title });

            for (const slide of input.slides) {
              presentation.addSlide({
                ...(slide.id !== undefined ? { id: slide.id } : {}),
                elements: slide.elements.map((element, index) => ({
                  type: "text",
                  text: element,
                  box: { x: 48, y: 48 + index * 32, width: 640, height: 24 },
                })),
              });
            }

            await exportPptx(presentation, { output: outputPath });
            const bytes = await readFile(outputPath);
            response.writeHead(200, {
              "content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
              "content-disposition": `attachment; filename="${example.id}.pptx"`,
              "content-length": bytes.byteLength,
            });
            response.end(bytes);
            return;
          } finally {
            await rm(temporaryDirectory, { recursive: true, force: true });
          }
        }

        const report = await buildExampleReport(example);
        response.writeHead(200, { "content-type": "application/json; charset=utf-8" });
        response.end(json(report));
        return;
      }

      if (url.pathname === "/api/health") {
        response.writeHead(200, { "content-type": "application/json; charset=utf-8" });
        response.end(json({ ok: true, examples: listExamples().length }));
        return;
      }

      if (url.pathname !== "/" && url.pathname !== "/index.html") {
        const assetPath = path.join(clientDir, url.pathname.replace(/^\/+/, ""));

        if (existsSync(assetPath)) {
          await serveFile(response, assetPath);
          return;
        }
      }

      await serveApp(response);
    } catch (error) {
      const status = error instanceof ExampleSourceError ? 400 : 500;
      response.writeHead(status, { "content-type": "application/json; charset=utf-8" });
      response.end(json({ error: String(error) }));
    }
  };
}

export function startWorkbenchServer(listenPort = port): Promise<Server> {
  const server = createWorkbenchServer();

  return new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(listenPort, () => {
      server.off("error", reject);
      resolve(server);
    });
  });
}

if (process.env.VITEST !== "true") {
  void startWorkbenchServer().then(() => {
    console.log(`PPTKit workbench running at http://localhost:${port}`);
  });
}
