import { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { ArrowDownToLine, ArrowLeft, ArrowRight, BookOpen, FileUp, Presentation } from "lucide-react";
import { defaultExample, examples, loadNativeSample } from "./examples";
import { download, inspectDocument, maxFileBytes, pptxMimeType, preview, previewImport, saveImport } from "./presentation";
import "./style.css";

const repository = "https://github.com/openHacking/pptkit";
const importExample = { id: "import", title: "PowerPoint import", description: "" };
const catalog = [defaultExample, ...examples, importExample];

type Loaded = {
  id: string;
  filename: string;
  preview: Awaited<ReturnType<typeof preview>>;
  imported?: Awaited<ReturnType<typeof previewImport>>;
};

const initialExample = catalog.find(({ id }) => `#${id}` === location.hash)?.id ?? defaultExample.id;

function previewPage(svg: string) {
  return `<!doctype html><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data: blob:; font-src data:; style-src 'unsafe-inline'"><style>html,body{margin:0;width:100%;height:100%;overflow:hidden}svg{display:block;width:100%;height:100%}</style>${svg}`;
}

function App() {
  const [loaded, setLoaded] = useState<Loaded>();
  const [slideIndex, setSlideIndex] = useState(0);
  const [source, setSource] = useState(defaultExample.source);
  const [appliedSource, setAppliedSource] = useState(defaultExample.source);
  const [replacement, setReplacement] = useState("");
  const [exportWarnings, setExportWarnings] = useState<readonly unknown[]>([]);
  const [pending, setPending] = useState("");
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const fileInput = useRef<HTMLInputElement>(null);
  const running = useRef(false);

  async function run(label: string, action: () => Promise<void>) {
    if (running.current) return;
    running.current = true;
    setPending(label);
    setError("");
    setMessage("");
    try {
      // Let the browser paint feedback before parsing or laying out a large file.
      await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
      await action();
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason));
    } finally {
      setPending("");
      running.current = false;
    }
  }

  function show(next: Loaded) {
    setLoaded(next);
    setSlideIndex(0);
    setReplacement(next.imported?.editableText ?? "");
    setExportWarnings([]);
    history.replaceState(null, "", `#${next.id}`);
  }

  async function openImport(bytes: Uint8Array, filename: string) {
    const imported = await previewImport(bytes);
    show({ id: importExample.id, filename, preview: imported, imported });
  }

  function openExample(id: string) {
    return run("Opening example…", async () => {
      if (id === importExample.id) {
        await openImport(await loadNativeSample(), "powerpoint-sample");
        return;
      }
      if (id === defaultExample.id) {
        show({ id, filename: id, preview: await preview(defaultExample.create()) });
        setSource(defaultExample.source);
        setAppliedSource(defaultExample.source);
        return;
      }
      const example = examples.find((item) => item.id === id);
      if (example) show({ id, filename: id, preview: await preview(example.create()) });
    });
  }

  useEffect(() => { void openExample(initialExample); }, []);

  const busy = Boolean(pending);
  const imported = loaded?.imported;
  const isFullDeck = loaded?.id === defaultExample.id;
  const dirty = isFullDeck && source !== appliedSource;
  const rendering = loaded?.preview.svg;
  const slide = rendering?.slides[slideIndex];
  const warnings = [...(rendering?.warnings ?? []), ...exportWarnings];
  const inspection = useMemo(() => loaded && (loaded.imported?.inspection ?? inspectDocument(loaded.preview.document)), [loaded]);
  const example = catalog.find(({ id }) => id === loaded?.id);
  const description = imported
    ? `${loaded.filename}.pptx · Partial content preview. See inspection results for import limitations.`
    : example?.description ?? "";

  return (
    <>
      <header className="site-header">
        <a className="brand" href={import.meta.env.BASE_URL}><Presentation aria-hidden="true" /> PPTKit <span>Playground</span></a>
        <nav aria-label="Project links">
          <a href={`${repository}/tree/main/docs`}><BookOpen aria-hidden="true" /> Docs</a>
          <a href={repository}>GitHub</a>
        </nav>
      </header>
      <main>
        <div className="intro">
          <h1>Explore PPTKit</h1>
          <p>Preview example decks, download editable PowerPoint files, or inspect a deck from your device.</p>
        </div>
        <div className="workspace">
          <aside aria-label="Example catalog">
            <h2>Examples</h2>
            <ul>
              {catalog.map(({ id, title }) => (
                <li key={id}><button aria-current={loaded?.id === id ? "true" : undefined} disabled={busy} onClick={() => void openExample(id)}>{title}</button></li>
              ))}
            </ul>
            <input ref={fileInput} type="file" hidden accept={`.pptx,${pptxMimeType}`} onChange={(event) => {
              const file = event.currentTarget.files?.[0];
              event.currentTarget.value = "";
              if (!file) return;
              void run("Importing PPTX…", async () => {
                if (file.size > maxFileBytes) throw new Error("Choose a PPTX file up to 32 MiB.");
                await openImport(new Uint8Array(await file.arrayBuffer()), file.name.replace(/\.pptx$/i, ""));
              });
            }} />
            <button className="upload" disabled={busy} onClick={() => fileInput.current?.click()}><FileUp aria-hidden="true" /> Open local PPTX</button>
            <p className="local-note">Files stay in your browser.</p>
          </aside>
          <section className="example" aria-busy={busy}>
            {loaded && rendering && (
              <>
                <div className="example-header">
                  <div><h2>{example?.title}</h2><p>{description}</p></div>
                  <button className="primary" disabled={busy || dirty} onClick={() => void run("Preparing PPTX…", async () => {
                    if (imported) {
                      const result = await saveImport(imported.bytes);
                      download(result.bytes.slice().buffer, `${loaded.filename}-preserved.pptx`, pptxMimeType);
                    } else {
                      const { generatePptx } = await import("@pptkit/pptx-exporter");
                      const result = await generatePptx(loaded.preview.document);
                      download(result.bytes.slice().buffer, `${loaded.filename}.pptx`, pptxMimeType);
                      setExportWarnings(result.warnings);
                    }
                    setMessage("PPTX downloaded.");
                  })}><ArrowDownToLine aria-hidden="true" /> {imported ? "Download preserved PPTX" : "Download PPTX"}</button>
                </div>
                {slide ? (
                  <>
                    <div className="stage">
                      <iframe title="Slide preview" sandbox="" tabIndex={-1} srcDoc={previewPage(slide.svg)} style={{ aspectRatio: `${rendering.width} / ${rendering.height}` }} />
                    </div>
                    <div className="slide-controls">
                      <div className="pagination">
                        <button aria-label="Previous slide" disabled={busy || slideIndex === 0} onClick={() => setSlideIndex(slideIndex - 1)}><ArrowLeft aria-hidden="true" /></button>
                        <span aria-live="polite">Slide {slideIndex + 1} of {rendering.slides.length}{slide.hidden ? " · hidden" : ""}</span>
                        <button aria-label="Next slide" disabled={busy || slideIndex + 1 >= rendering.slides.length} onClick={() => setSlideIndex(slideIndex + 1)}><ArrowRight aria-hidden="true" /></button>
                      </div>
                      <button disabled={busy || dirty} onClick={() => download(slide.svg, `${loaded.filename}-slide-${slideIndex + 1}.svg`, "image/svg+xml")}><ArrowDownToLine aria-hidden="true" /> SVG</button>
                    </div>
                  </>
                ) : <p className="empty">This presentation has no slides to preview.</p>}
                {isFullDeck && (
                  <details>
                    <summary>Edit example JSON</summary>
                    <textarea aria-label="Example JSON" value={source} disabled={busy} spellCheck={false} onChange={(event) => setSource(event.target.value)} />
                    <div className="actions">
                      <button className="primary" disabled={busy || !dirty} onClick={() => void run("Updating preview…", async () => {
                        const updated = await preview(defaultExample.create(source));
                        setLoaded({ ...loaded, preview: updated });
                        setAppliedSource(source);
                        setSlideIndex(Math.min(slideIndex, Math.max(0, updated.svg.slides.length - 1)));
                        setExportWarnings([]);
                        setMessage("Preview updated.");
                      })}>Apply changes</button>
                      <button disabled={busy || !dirty} onClick={() => { setSource(appliedSource); setError(""); }}>Reset changes</button>
                    </div>
                  </details>
                )}
                {imported?.editableText !== undefined && (
                  <details>
                    <summary>Edit a preserved text run</summary>
                    <label>Text from the first supported text box
                      <input value={replacement} disabled={busy} onChange={(event) => setReplacement(event.target.value)} />
                    </label>
                    <button className="primary" disabled={busy || !replacement || replacement === imported.editableText} onClick={() => void run("Checking preservation save…", async () => {
                      const saved = await saveImport(imported.bytes, replacement);
                      const filename = `${loaded.filename}-edited`;
                      await openImport(saved.bytes, filename);
                      download(saved.bytes.slice().buffer, `${filename}.pptx`, pptxMimeType);
                      setMessage("Text updated and PPTX downloaded.");
                    })}>Save edited PPTX</button>
                  </details>
                )}
                <details>
                  <summary>{imported ? "Import inspection" : "Validation & layout"}{warnings.length > 0 ? ` · ${warnings.length} warnings` : ""}</summary>
                  <pre>{JSON.stringify({ ...inspection, warnings }, null, 2)}</pre>
                </details>
              </>
            )}
            {pending && <p className="feedback" role="status">{pending}</p>}
            {message && <p className="feedback" role="status">{message}</p>}
            {error && <p className="feedback error" role="alert">{error}</p>}
          </section>
        </div>
        <footer><a href={`${repository}/blob/main/examples/README.md#online-playground`}>More examples and Node.js workflows</a></footer>
      </main>
    </>
  );
}

createRoot(document.getElementById("root")!).render(<App />);
