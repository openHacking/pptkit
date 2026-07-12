type FeatureId = string;

interface ExampleSummary {
  id: string;
  title: string;
  feature: FeatureId;
  description: string;
  inputKind: string;
  scenarioTags: string[];
  status: string;
}

interface WorkbenchPayload {
  features: FeatureId[];
  examples: ExampleSummary[];
}

interface ExampleReport {
  example: {
    id: string;
    title: string;
    feature: FeatureId;
    description: string;
    inputKind: string;
    source: {
      label: string;
      content: string;
    };
    scenarioTags: string[];
    expectedCapabilities: {
      normalize: string;
      render: string;
      exportPptx: string;
    };
    status: string;
  };
  normalizedDocument: {
    title: string;
    summary: string;
    slideCount: number;
    slides: Array<{
      id: string;
      title: string;
      elements: string[];
    }>;
  };
  visualPreview: {
    status: string;
    slides: Array<{
      id: string;
      title: string;
      body: string[];
    }>;
  };
  renderResult: {
    slideCount: number;
    status: string;
  };
  exportResult: {
    output: string;
    slideCount: number;
    status: string;
  };
  diagnostics: string[];
}

const appRoot = document.querySelector("#app");

if (!(appRoot instanceof HTMLElement)) {
  throw new Error("Missing #app root.");
}

const rootElement: HTMLElement = appRoot;

function escapeHtml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function capabilityBadge(value: string): string {
  const tone =
    value === "implemented"
      ? "success"
      : value === "placeholder"
        ? "warning"
        : "muted";

  return `<span class="badge badge-${tone}">${escapeHtml(value)}</span>`;
}

async function fetchJson<T>(url: string): Promise<T> {
  const response = await fetch(url);

  if (!response.ok) {
    throw new Error(`Request failed: ${response.status}`);
  }

  return (await response.json()) as T;
}

function buildFeatureNav(features: FeatureId[], activeFeature: FeatureId): string {
  return features
    .map(
      (feature) => `
        <button class="feature-tab${feature === activeFeature ? " is-active" : ""}" data-feature="${feature}">
          ${escapeHtml(feature)}
        </button>
      `,
    )
    .join("");
}

function buildExampleList(examples: ExampleSummary[], activeId: string): string {
  return examples
    .map(
      (example) => `
        <button class="case-item${example.id === activeId ? " is-active" : ""}" data-example-id="${example.id}">
          <span class="case-title">${escapeHtml(example.title)}</span>
          <span class="case-meta">${escapeHtml(example.inputKind)} • ${escapeHtml(example.status)}</span>
        </button>
      `,
    )
    .join("");
}

function buildSlides(report: ExampleReport): string {
  return report.visualPreview.slides
    .map(
      (slide, index) => `
        <article class="slide-card">
          <div class="slide-eyebrow">Slide ${index + 1}</div>
          <h4>${escapeHtml(slide.title)}</h4>
          <ul>
            ${slide.body.map((line) => `<li>${escapeHtml(line)}</li>`).join("")}
          </ul>
        </article>
      `,
    )
    .join("");
}

function renderWorkbench(
  payload: WorkbenchPayload,
  activeFeature: FeatureId,
  activeExamples: ExampleSummary[],
  report: ExampleReport,
) {
  rootElement.innerHTML = `
    <div class="shell">
      <header class="hero">
        <div>
          <p class="eyebrow">PPTKit Developer Workbench</p>
          <h1>Switch cases by feature to quickly inspect input, IR, preview, and export status</h1>
          <p class="hero-copy">This workbench is for internal development, prioritizing preview / export edge validation over final visual fidelity.</p>
        </div>
      </header>

      <nav class="feature-nav">${buildFeatureNav(payload.features, activeFeature)}</nav>

      <main class="layout">
        <aside class="panel panel-list">
          <div class="panel-header">
            <h2>Feature Cases</h2>
            <span>${activeExamples.length} cases</span>
          </div>
          <div class="case-list">${buildExampleList(activeExamples, report.example.id)}</div>
        </aside>

        <section class="panel panel-detail">
          <div class="panel-header">
            <div>
              <p class="eyebrow">${escapeHtml(report.example.feature)}</p>
              <h2>${escapeHtml(report.example.title)}</h2>
            </div>
            <span class="status-pill">${escapeHtml(report.example.status)}</span>
          </div>

          <p class="description">${escapeHtml(report.example.description)}</p>

          <div class="meta-row">
            <div><strong>Input</strong> ${escapeHtml(report.example.inputKind)}</div>
            <div><strong>Scenarios</strong> ${escapeHtml(report.example.scenarioTags.join(", "))}</div>
          </div>

          <div class="capability-grid">
            <div class="capability-card"><span>Normalize</span>${capabilityBadge(report.example.expectedCapabilities.normalize)}</div>
            <div class="capability-card"><span>Render</span>${capabilityBadge(report.example.expectedCapabilities.render)}</div>
            <div class="capability-card"><span>Export PPTX</span>${capabilityBadge(report.example.expectedCapabilities.exportPptx)}</div>
          </div>

          <div class="detail-grid">
            <section class="subpanel">
              <h3>Source</h3>
              <p class="subtle">${escapeHtml(report.example.source.label)}</p>
              <pre>${escapeHtml(report.example.source.content)}</pre>
            </section>

            <section class="subpanel">
              <h3>Normalized Document</h3>
              <p class="subtle">${escapeHtml(report.normalizedDocument.summary)}</p>
              <pre>${escapeHtml(JSON.stringify(report.normalizedDocument, null, 2))}</pre>
            </section>

            <section class="subpanel preview-panel">
              <div class="preview-header">
                <h3>Structural Preview</h3>
                <span>${escapeHtml(report.visualPreview.status)}</span>
              </div>
              <div class="slide-grid">${buildSlides(report)}</div>
            </section>

            <section class="subpanel">
              <h3>Render / Export Results</h3>
              <pre>${escapeHtml(
                JSON.stringify(
                  {
                    render: report.renderResult,
                    exportPptx: report.exportResult,
                  },
                  null,
                  2,
                ),
              )}</pre>
            </section>

            <section class="subpanel">
              <h3>Diagnostics</h3>
              <ul class="diagnostics">
                ${report.diagnostics.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}
              </ul>
            </section>
          </div>
        </section>
      </main>
    </div>
  `;
}

async function main() {
  const payload = await fetchJson<WorkbenchPayload>("/api/workbench");
  const initialFeature = payload.features[0];

  if (initialFeature === undefined) {
    throw new Error("No features registered.");
  }

  let activeFeature: FeatureId = initialFeature;

  let activeExamples = payload.examples.filter((example) => example.feature === activeFeature);
  let activeExampleId = activeExamples[0]?.id;

  if (activeExampleId === undefined) {
    throw new Error(`No examples registered for feature ${activeFeature}.`);
  }

  async function syncView(nextFeature?: FeatureId, nextExampleId?: string) {
    if (nextFeature !== undefined) {
      activeFeature = nextFeature;
      activeExamples = payload.examples.filter((example) => example.feature === activeFeature);
      activeExampleId = activeExamples[0]?.id;
    }

    if (nextExampleId !== undefined) {
      activeExampleId = nextExampleId;
    }

    if (activeExampleId === undefined) {
      throw new Error(`No examples registered for feature ${activeFeature}.`);
    }

    const report = await fetchJson<ExampleReport>(`/api/examples/${activeExampleId}`);
    renderWorkbench(payload, activeFeature, activeExamples, report);
    wireEvents();
  }

  function wireEvents() {
    for (const tab of Array.from(document.querySelectorAll<HTMLElement>("[data-feature]"))) {
      tab.onclick = () => {
        const feature = tab.dataset.feature as FeatureId | undefined;

        if (feature !== undefined && feature !== activeFeature) {
          void syncView(feature);
        }
      };
    }

    for (const item of Array.from(
      document.querySelectorAll<HTMLElement>("[data-example-id]"),
    )) {
      item.onclick = () => {
        const exampleId = item.dataset.exampleId;

        if (exampleId !== undefined && exampleId !== activeExampleId) {
          void syncView(undefined, exampleId);
        }
      };
    }
  }

  await syncView();
}

void main().catch((error: unknown) => {
  rootElement.innerHTML = `<pre class="error">${escapeHtml(String(error))}</pre>`;
});
