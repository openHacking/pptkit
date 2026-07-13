import { AlertCircle, Check, Download, LoaderCircle, Presentation, RotateCcw, TriangleAlert } from "lucide-react";
import { useEffect, useState } from "react";
import { generatePptx } from "@pptkit/pptx-exporter";
import type { CapabilityStatus, ExampleReport, ExampleSummary, FeatureId, WorkbenchPayload } from "../example-types";
import { createExamplePresentation } from "../presentation-builder";
import { parseExampleSource } from "../source-parser";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";

async function fetchJson<T>(url: string): Promise<T> {
  const response = await fetch(url);

  if (!response.ok) {
    throw new Error(`Request failed: ${response.status}`);
  }

  return (await response.json()) as T;
}

async function postJson<T>(url: string, body: unknown): Promise<T> {
  const response = await fetch(url, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    const payload = (await response.json().catch(() => null)) as { error?: string } | null;
    throw new Error(payload?.error ?? `Request failed: ${response.status}`);
  }

  return (await response.json()) as T;
}

function capabilityVariant(value: CapabilityStatus): "default" | "secondary" | "outline" {
  if (value === "implemented") {
    return "default";
  }

  if (value === "placeholder") {
    return "secondary";
  }

  return "outline";
}

function statusVariant(value: string): "default" | "secondary" | "outline" {
  if (value === "ready" || value === "implemented") {
    return "default";
  }

  if (value === "placeholder") {
    return "secondary";
  }

  return "outline";
}

function JsonBlock({ value }: { value: unknown }) {
  return (
    <ScrollArea className="h-[28rem] rounded-md border">
      <pre className="p-4 text-xs leading-5 whitespace-pre-wrap">
        {JSON.stringify(value, null, 2)}
      </pre>
    </ScrollArea>
  );
}

function SourceEditor({
  label,
  value,
  dirty,
  error,
  applying,
  browserExporting,
  serverExporting,
  onChange,
  onApply,
  onReset,
  onBrowserExport,
  onServerExport,
}: {
  label: string;
  value: string;
  dirty: boolean;
  error: string | null;
  applying: boolean;
  browserExporting: boolean;
  serverExporting: boolean;
  onChange: (value: string) => void;
  onApply: () => void;
  onReset: () => void;
  onBrowserExport: () => void;
  onServerExport: () => void;
}) {
  const busy = applying || browserExporting || serverExporting;

  return (
    <div className="space-y-2">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div>
          <p className="text-muted-foreground text-sm">{label}</p>
          <p className="text-muted-foreground mt-1 text-xs">
            {dirty ? "Unsaved changes" : "Applied source"}
          </p>
        </div>
        <div className="flex flex-wrap gap-2">
          <Button type="button" variant="outline" size="sm" onClick={onReset} disabled={!dirty || busy}>
            <RotateCcw />
            Reset
          </Button>
          <Button type="button" size="sm" onClick={onApply} disabled={!dirty || busy}>
            {applying ? <LoaderCircle className="animate-spin" /> : <Check />}
            {applying ? "Applying..." : "Apply changes"}
          </Button>
          <Button type="button" variant="secondary" size="sm" onClick={onBrowserExport} disabled={dirty || busy}>
            {browserExporting ? <LoaderCircle className="animate-spin" /> : <Download />}
            {browserExporting ? "Exporting in browser..." : "Export in browser"}
          </Button>
          <Button type="button" variant="outline" size="sm" onClick={onServerExport} disabled={dirty || busy}>
            {serverExporting ? <LoaderCircle className="animate-spin" /> : <Download />}
            {serverExporting ? "Exporting via server..." : "Export via server"}
          </Button>
        </div>
      </div>
      <textarea
        aria-label={label}
        className="min-h-[28rem] w-full resize-y rounded-md border bg-muted/20 p-4 font-mono text-xs leading-5 outline-none transition focus-visible:ring-2 focus-visible:ring-ring"
        value={value}
        onChange={(event) => onChange(event.target.value)}
        spellCheck={false}
      />
      {error !== null && <p className="text-destructive text-sm">{error}</p>}
    </div>
  );
}

function FeatureTabs({
  features,
  activeFeature,
  onFeatureChange,
}: {
  features: FeatureId[];
  activeFeature: FeatureId;
  onFeatureChange: (feature: FeatureId) => void;
}) {
  return (
    <Tabs value={activeFeature} onValueChange={(value) => onFeatureChange(value as FeatureId)} className="gap-3">
      <ScrollArea className="w-full whitespace-nowrap">
        <TabsList>
          {features.map((feature) => (
            <TabsTrigger key={feature} value={feature}>
              {feature}
            </TabsTrigger>
          ))}
        </TabsList>
      </ScrollArea>
    </Tabs>
  );
}

function ExampleList({
  examples,
  activeExampleId,
  onSelect,
}: {
  examples: ExampleSummary[];
  activeExampleId: string;
  onSelect: (exampleId: string) => void;
}) {
  return (
    <div className="space-y-2">
      {examples.map((example) => (
        <Button
          key={example.id}
          type="button"
          variant={example.id === activeExampleId ? "default" : "outline"}
          className="h-auto w-full items-start justify-start rounded-lg px-4 py-3 text-left"
          onClick={() => onSelect(example.id)}
        >
          <span className="font-medium">{example.title}</span>
        </Button>
      ))}
    </div>
  );
}

function ReportView({
  report,
  source,
  sourceDirty,
  sourceError,
  applying,
  browserExporting,
  serverExporting,
  exportMessage,
  onSourceChange,
  onApply,
  onReset,
  onBrowserExport,
  onServerExport,
}: {
  report: ExampleReport;
  source: string;
  sourceDirty: boolean;
  sourceError: string | null;
  applying: boolean;
  browserExporting: boolean;
  serverExporting: boolean;
  exportMessage: string | null;
  onSourceChange: (value: string) => void;
  onApply: () => void;
  onReset: () => void;
  onBrowserExport: () => void;
  onServerExport: () => void;
}) {
  return (
    <div className="flex h-full flex-col gap-5">
      <div className="flex flex-col gap-3 border-b pb-4 sm:flex-row sm:items-start sm:justify-between">
        <div>
          <h2 className="sr-only">{report.example.title}</h2>
          <p className="text-muted-foreground text-sm">{report.example.description}</p>
        </div>
        <Badge variant={statusVariant(report.example.status)}>{report.example.status}</Badge>
      </div>

      <div className="grid gap-4 md:grid-cols-3">
        <Card className="gap-3 py-5">
          <CardHeader>
            <CardTitle className="text-base">Normalize</CardTitle>
          </CardHeader>
          <CardContent>
            <Badge variant={capabilityVariant(report.example.expectedCapabilities.normalize)}>
              {report.example.expectedCapabilities.normalize}
            </Badge>
          </CardContent>
        </Card>
        <Card className="gap-3 py-5">
          <CardHeader>
            <CardTitle className="text-base">Render</CardTitle>
          </CardHeader>
          <CardContent>
            <Badge variant={capabilityVariant(report.example.expectedCapabilities.render)}>
              {report.example.expectedCapabilities.render}
            </Badge>
          </CardContent>
        </Card>
        <Card className="gap-3 py-5">
          <CardHeader>
            <CardTitle className="text-base">Export PPTX</CardTitle>
          </CardHeader>
          <CardContent>
            <Badge variant={capabilityVariant(report.example.expectedCapabilities.exportPptx)}>
              {report.example.expectedCapabilities.exportPptx}
            </Badge>
          </CardContent>
        </Card>
      </div>

      <Card className="gap-4 py-5">
        <CardHeader>
          <CardTitle>Structural Preview</CardTitle>
        </CardHeader>
        <CardContent className="grid gap-4 md:grid-cols-2 xl:grid-cols-3">
          {report.visualPreview.slides.map((slide, index) => (
            <Card key={slide.id} className="gap-3 py-5">
              <CardHeader>
                <div className="text-muted-foreground text-sm">Slide {index + 1}</div>
                <CardTitle className="text-base">{slide.title}</CardTitle>
              </CardHeader>
              <CardContent>
                <ul className="space-y-2 text-sm">
                  {slide.body.map((line) => (
                    <li key={line}>{line}</li>
                  ))}
                </ul>
              </CardContent>
            </Card>
          ))}
        </CardContent>
      </Card>

      <Tabs defaultValue="source" className="flex min-h-0 flex-1 flex-col gap-4">
        <TabsList>
          <TabsTrigger value="source">Source</TabsTrigger>
          <TabsTrigger value="normalized">Normalized</TabsTrigger>
          <TabsTrigger value="results">Results</TabsTrigger>
        </TabsList>
        <TabsContent value="source" className="min-h-0 flex-1">
          <Card className="h-full gap-4 py-5">
            <CardHeader>
              <CardTitle>Source</CardTitle>
            </CardHeader>
            <CardContent>
              <SourceEditor
                label={report.example.source.label}
                value={source}
                dirty={sourceDirty}
                error={sourceError}
                applying={applying}
                browserExporting={browserExporting}
                serverExporting={serverExporting}
                onChange={onSourceChange}
                onApply={onApply}
                onReset={onReset}
                onBrowserExport={onBrowserExport}
                onServerExport={onServerExport}
              />
              {exportMessage !== null && (
                <p className="flex items-center gap-2 text-sm text-emerald-700">
                  <Check className="size-4" />
                  {exportMessage}
                </p>
              )}
            </CardContent>
          </Card>
        </TabsContent>
        <TabsContent value="normalized" className="min-h-0 flex-1">
          <Card className="h-full gap-4 py-5">
            <CardHeader>
              <CardTitle>Normalized Document</CardTitle>
            </CardHeader>
            <CardContent>
              <JsonBlock value={report.normalizedDocument} />
            </CardContent>
          </Card>
        </TabsContent>
        <TabsContent value="results" className="min-h-0 flex-1">
          <Card className="h-full gap-4 py-5">
            <CardHeader>
              <CardTitle>Results</CardTitle>
            </CardHeader>
            <CardContent>
              <JsonBlock
                value={{
                  render: report.renderResult,
                  exportPptx: report.exportResult,
                }}
              />
            </CardContent>
          </Card>
        </TabsContent>
      </Tabs>

      <Alert>
        <TriangleAlert className="size-4" />
        <AlertTitle>Diagnostics</AlertTitle>
        <AlertDescription>
          <ul className="space-y-2">
            {report.diagnostics.map((item) => (
              <li key={item}>{item}</li>
            ))}
          </ul>
        </AlertDescription>
      </Alert>
    </div>
  );
}

export default function App() {
  const [payload, setPayload] = useState<WorkbenchPayload | null>(null);
  const [activeFeature, setActiveFeature] = useState<FeatureId | null>(null);
  const [activeExampleId, setActiveExampleId] = useState<string | null>(null);
  const [report, setReport] = useState<ExampleReport | null>(null);
  const [source, setSource] = useState("");
  const [appliedSource, setAppliedSource] = useState("");
  const [sourceError, setSourceError] = useState<string | null>(null);
  const [applying, setApplying] = useState(false);
  const [browserExporting, setBrowserExporting] = useState(false);
  const [serverExporting, setServerExporting] = useState(false);
  const [exportMessage, setExportMessage] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let cancelled = false;

    async function loadWorkbench() {
      try {
        setLoading(true);
        const nextPayload = await fetchJson<WorkbenchPayload>("/api/workbench");
        const initialFeature = nextPayload.features[0];

        if (initialFeature === undefined) {
          throw new Error("No features registered.");
        }

        const initialExample = nextPayload.examples.find((example) => example.feature === initialFeature);

        if (initialExample === undefined) {
          throw new Error(`No examples registered for feature ${initialFeature}.`);
        }

        const nextReport = await fetchJson<ExampleReport>(`/api/examples/${initialExample.id}`);

        if (cancelled) {
          return;
        }

        setPayload(nextPayload);
        setActiveFeature(initialFeature);
        setActiveExampleId(initialExample.id);
        setReport(nextReport);
        setSource(nextReport.example.source.content);
        setAppliedSource(nextReport.example.source.content);
        setError(null);
      } catch (nextError) {
        if (!cancelled) {
          setError(String(nextError));
        }
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
    }

    void loadWorkbench();

    return () => {
      cancelled = true;
    };
  }, []);

  async function selectFeature(feature: FeatureId) {
    if (payload === null || feature === activeFeature) {
      return;
    }

    const nextExample = payload.examples.find((example) => example.feature === feature);

    if (nextExample === undefined) {
      setError(`No examples registered for feature ${feature}.`);
      return;
    }

    setLoading(true);

    try {
      const nextReport = await fetchJson<ExampleReport>(`/api/examples/${nextExample.id}`);
      setActiveFeature(feature);
      setActiveExampleId(nextExample.id);
      setReport(nextReport);
      setSource(nextReport.example.source.content);
      setAppliedSource(nextReport.example.source.content);
      setSourceError(null);
      setExportMessage(null);
      setError(null);
    } catch (nextError) {
      setError(String(nextError));
    } finally {
      setLoading(false);
    }
  }

  async function selectExample(exampleId: string) {
    if (exampleId === activeExampleId) {
      return;
    }

    setLoading(true);

    try {
      const nextReport = await fetchJson<ExampleReport>(`/api/examples/${exampleId}`);
      setActiveExampleId(exampleId);
      setReport(nextReport);
      setSource(nextReport.example.source.content);
      setAppliedSource(nextReport.example.source.content);
      setSourceError(null);
      setExportMessage(null);
      setError(null);
    } catch (nextError) {
      setError(String(nextError));
    } finally {
      setLoading(false);
    }
  }

  async function applySource() {
    if (report === null || source === appliedSource) {
      return;
    }

    setApplying(true);
    setSourceError(null);
    setExportMessage(null);

    try {
      const nextReport = await postJson<ExampleReport>(`/api/examples/${report.example.id}/report`, { source });
      setReport(nextReport);
      setAppliedSource(source);
      setSourceError(null);
    } catch (nextError) {
      setSourceError(String(nextError).replace(/^Error: /, ""));
    } finally {
      setApplying(false);
    }
  }

  function resetSource() {
    setSource(appliedSource);
    setSourceError(null);
    setExportMessage(null);
  }

  function downloadPptx(bytes: Uint8Array, filename: string) {
    const blob = new Blob([bytes.slice().buffer], {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });
    const downloadUrl = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = downloadUrl;
    link.download = filename;
    link.click();
    URL.revokeObjectURL(downloadUrl);
  }

  async function exportInBrowser() {
    if (report === null || source !== appliedSource) {
      return;
    }

    setBrowserExporting(true);
    setSourceError(null);
    setExportMessage(null);

    try {
      const presentation = createExamplePresentation(parseExampleSource(appliedSource));
      const result = await generatePptx(presentation);
      downloadPptx(result.bytes, `${report.example.id}.pptx`);
      setExportMessage("PPTX exported in browser successfully.");
    } catch (nextError) {
      setSourceError(String(nextError).replace(/^Error: /, ""));
    } finally {
      setBrowserExporting(false);
    }
  }

  async function exportViaServer() {
    if (report === null || source !== appliedSource) {
      return;
    }

    setServerExporting(true);
    setSourceError(null);
    setExportMessage(null);

    try {
      const response = await fetch(`/api/examples/${report.example.id}/export`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ source: appliedSource }),
      });

      if (!response.ok) {
        const payload = (await response.json().catch(() => null)) as { error?: string } | null;
        throw new Error(payload?.error ?? `Export failed: ${response.status}`);
      }

      const bytes = new Uint8Array(await response.arrayBuffer());
      downloadPptx(bytes, `${report.example.id}.pptx`);
      setExportMessage("PPTX exported via server successfully.");
    } catch (nextError) {
      setSourceError(String(nextError).replace(/^Error: /, ""));
    } finally {
      setServerExporting(false);
    }
  }

  if (error !== null) {
    return (
      <main className="flex min-h-screen items-start p-4 md:p-6">
        <Alert variant="destructive">
          <AlertCircle className="size-4" />
          <AlertTitle>Workbench failed to load</AlertTitle>
          <AlertDescription>{error}</AlertDescription>
        </Alert>
      </main>
    );
  }

  if (payload === null || activeFeature === null || activeExampleId === null || report === null) {
    return (
      <main className="flex min-h-screen items-center justify-center p-4 md:p-6">
        <div className="flex items-center gap-2 text-sm text-muted-foreground">
          <LoaderCircle className="size-4 animate-spin" />
          Loading...
        </div>
      </main>
    );
  }

  const activeExamples = payload.examples.filter((example) => example.feature === activeFeature);

  return (
    <main className="flex min-h-screen flex-col gap-4 p-4 md:p-6">
      <header className="flex items-center gap-2 border-b pb-3 text-sm font-medium">
        <Presentation className="size-4" />
        PPTKit Workbench
      </header>
      <FeatureTabs features={payload.features} activeFeature={activeFeature} onFeatureChange={selectFeature} />

      <div className="grid min-h-0 flex-1 gap-4 lg:grid-cols-[280px_minmax(0,1fr)]">
        <Card className="min-h-0 gap-4 py-5">
          <CardHeader>
            <CardTitle>Feature Cases</CardTitle>
          </CardHeader>
          <CardContent className="min-h-0 flex-1">
            <ScrollArea className="h-full pr-3">
              <ExampleList
                examples={activeExamples}
                activeExampleId={activeExampleId}
                onSelect={selectExample}
              />
            </ScrollArea>
          </CardContent>
        </Card>

        <div className="flex min-h-0 flex-col gap-4">
          {loading && (
            <Alert>
              <LoaderCircle className="size-4 animate-spin" />
              <AlertTitle>Loading</AlertTitle>
            </Alert>
          )}

          <ScrollArea className="min-h-0 flex-1 rounded-xl border">
            <div className="p-5">
              <ReportView
                report={report}
                source={source}
                sourceDirty={source !== appliedSource}
                sourceError={sourceError}
                applying={applying}
                browserExporting={browserExporting}
                serverExporting={serverExporting}
                exportMessage={exportMessage}
                onSourceChange={(value) => {
                  setSource(value);
                  setSourceError(null);
                  setExportMessage(null);
                }}
                onApply={() => void applySource()}
                onReset={resetSource}
                onBrowserExport={() => void exportInBrowser()}
                onServerExport={() => void exportViaServer()}
              />
            </div>
          </ScrollArea>
        </div>
      </div>
    </main>
  );
}
