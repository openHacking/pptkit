import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { getExampleById, listExamples, listExampleSummaries, listFeatures } from "./example-registry.js";
import { buildExampleReport } from "./report-builder.js";

const distDir = path.dirname(fileURLToPath(import.meta.url));
const port = Number(process.env.PORT ?? "3210");

function json(data: unknown) {
  return JSON.stringify(data, null, 2);
}

function html() {
  return `<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>PPTKit Developer Workbench</title>
    <style>
      :root {
        color-scheme: light;
        --bg: #f3efe6;
        --paper: #fffdf8;
        --ink: #182126;
        --muted: #66737c;
        --line: #d6cfc1;
        --accent: #116466;
        --accent-soft: #d6ece4;
        --warning: #ad6f00;
        --warning-soft: #fff2cf;
        --success: #0f766e;
        --success-soft: #daf5ef;
        --shadow: 0 18px 45px rgba(24, 33, 38, 0.08);
      }

      * {
        box-sizing: border-box;
      }

      body {
        margin: 0;
        font-family: "Iowan Old Style", "Palatino Linotype", "Book Antiqua", serif;
        color: var(--ink);
        background:
          radial-gradient(circle at top left, rgba(17, 100, 102, 0.12), transparent 32%),
          linear-gradient(180deg, #f7f3ea 0%, #efe7d6 100%);
      }

      button,
      pre,
      code {
        font-family: "SFMono-Regular", "Menlo", monospace;
      }

      button {
        border: 0;
        background: transparent;
        color: inherit;
        cursor: pointer;
      }

      .shell {
        padding: 32px;
        min-height: 100vh;
      }

      .hero,
      .feature-nav,
      .panel {
        background: rgba(255, 253, 248, 0.92);
        backdrop-filter: blur(8px);
        border: 1px solid rgba(214, 207, 193, 0.9);
        box-shadow: var(--shadow);
      }

      .hero {
        padding: 28px 30px;
        border-radius: 28px;
        margin-bottom: 18px;
      }

      .hero h1 {
        margin: 6px 0 10px;
        font-size: clamp(28px, 4vw, 44px);
        line-height: 1.05;
        max-width: 13ch;
      }

      .hero-copy,
      .description,
      .subtle,
      .case-meta,
      .diagnostics,
      .meta-row,
      .panel-header span {
        color: var(--muted);
      }

      .eyebrow,
      .slide-eyebrow {
        margin: 0;
        text-transform: uppercase;
        letter-spacing: 0.12em;
        font-size: 12px;
        color: var(--accent);
      }

      .feature-nav {
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        padding: 12px;
        border-radius: 20px;
        margin-bottom: 18px;
      }

      .feature-tab,
      .case-item,
      .status-pill,
      .badge {
        border-radius: 999px;
      }

      .feature-tab {
        padding: 10px 14px;
        background: #f8f4eb;
        border: 1px solid transparent;
      }

      .feature-tab.is-active {
        background: var(--accent);
        color: #fff;
      }

      .layout {
        display: grid;
        grid-template-columns: minmax(240px, 280px) minmax(0, 1fr);
        gap: 18px;
      }

      .panel {
        border-radius: 28px;
        padding: 22px;
      }

      .panel-header,
      .preview-header,
      .meta-row,
      .capability-card,
      .case-item {
        display: flex;
        justify-content: space-between;
        gap: 12px;
      }

      .panel-header {
        align-items: start;
      }

      .panel-header h2,
      .subpanel h3,
      .slide-card h4 {
        margin: 0;
      }

      .case-list {
        display: grid;
        gap: 10px;
        margin-top: 18px;
      }

      .case-item {
        text-align: left;
        flex-direction: column;
        align-items: start;
        padding: 14px 16px;
        background: #f8f4eb;
        border: 1px solid transparent;
      }

      .case-item.is-active {
        border-color: rgba(17, 100, 102, 0.28);
        background: var(--accent-soft);
      }

      .case-title {
        font-weight: 700;
        font-size: 15px;
      }

      .status-pill {
        padding: 8px 12px;
        background: #f8f4eb;
        font-size: 12px;
        text-transform: uppercase;
      }

      .meta-row {
        display: flex;
        flex-wrap: wrap;
        margin: 18px 0;
      }

      .capability-grid,
      .detail-grid,
      .slide-grid {
        display: grid;
        gap: 14px;
      }

      .capability-grid {
        grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
        margin-bottom: 18px;
      }

      .capability-card,
      .subpanel,
      .slide-card {
        border: 1px solid var(--line);
        background: #fffdf8;
      }

      .capability-card {
        padding: 14px 16px;
        border-radius: 18px;
        align-items: center;
      }

      .detail-grid {
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }

      .subpanel {
        border-radius: 22px;
        padding: 18px;
      }

      .preview-panel {
        grid-column: 1 / -1;
      }

      .slide-grid {
        grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      }

      .slide-card {
        border-radius: 24px;
        padding: 18px;
        background:
          linear-gradient(180deg, rgba(17, 100, 102, 0.08), transparent 42%),
          #fffdf8;
      }

      .slide-card ul,
      .diagnostics {
        margin: 12px 0 0;
        padding-left: 18px;
      }

      .badge {
        display: inline-flex;
        align-items: center;
        padding: 5px 9px;
        font-size: 12px;
      }

      .badge-success {
        background: var(--success-soft);
        color: var(--success);
      }

      .badge-warning {
        background: var(--warning-soft);
        color: var(--warning);
      }

      .badge-muted {
        background: #ece7de;
        color: #64584c;
      }

      pre {
        margin: 12px 0 0;
        padding: 14px;
        border-radius: 16px;
        overflow: auto;
        white-space: pre-wrap;
        background: #f5f1e8;
        border: 1px solid var(--line);
        font-size: 12px;
      }

      .error {
        margin: 32px;
      }

      @media (max-width: 920px) {
        .shell {
          padding: 18px;
        }

        .layout,
        .detail-grid {
          grid-template-columns: 1fr;
        }
      }
    </style>
  </head>
  <body>
    <div id="app"></div>
    <script type="module" src="/client.js"></script>
  </body>
</html>`;
}

async function serveAsset(assetPath: string) {
  const filePath = path.join(distDir, assetPath);
  const content = await readFile(filePath, "utf8");
  const contentType = assetPath.endsWith(".js")
    ? "text/javascript; charset=utf-8"
    : "application/octet-stream";

  return {
    statusCode: 200,
    headers: { "content-type": contentType },
    body: content,
  };
}

const server = createServer(async (request, response) => {
  try {
    const url = new URL(request.url ?? "/", `http://${request.headers.host ?? "localhost"}`);

    if (url.pathname === "/") {
      response.writeHead(200, { "content-type": "text/html; charset=utf-8" });
      response.end(html());
      return;
    }

    if (url.pathname === "/client.js") {
      const asset = await serveAsset("client.js");
      response.writeHead(asset.statusCode, asset.headers);
      response.end(asset.body);
      return;
    }

    if (url.pathname === "/api/workbench") {
      response.writeHead(200, { "content-type": "application/json; charset=utf-8" });
      response.end(
        json({
          features: listFeatures(),
          examples: listExampleSummaries(),
        }),
      );
      return;
    }

    if (url.pathname.startsWith("/api/examples/")) {
      const id = url.pathname.replace("/api/examples/", "");
      const example = getExampleById(id);

      if (example === undefined) {
        response.writeHead(404, { "content-type": "application/json; charset=utf-8" });
        response.end(json({ error: `Unknown example: ${id}` }));
        return;
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

    response.writeHead(404, { "content-type": "application/json; charset=utf-8" });
    response.end(json({ error: "Not found" }));
  } catch (error) {
    response.writeHead(500, { "content-type": "application/json; charset=utf-8" });
    response.end(json({ error: String(error) }));
  }
});

server.listen(port, () => {
  console.log(`PPTKit workbench running at http://localhost:${port}`);
});
