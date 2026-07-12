import { AlertCircle, LoaderCircle, Presentation, TriangleAlert } from "lucide-react";
import { useEffect, useState } from "react";
import type { CapabilityStatus, ExampleReport, ExampleSummary, FeatureId, WorkbenchPayload } from "../example-types";
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

function SourceBlock({ label, content }: { label: string; content: string }) {
  return (
    <div className="space-y-2">
      <p className="text-muted-foreground text-sm">{label}</p>
      <ScrollArea className="h-[28rem] rounded-md border">
        <pre className="p-4 text-xs leading-5 whitespace-pre-wrap">{content}</pre>
      </ScrollArea>
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

function ReportView({ report }: { report: ExampleReport }) {
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
              <SourceBlock label={report.example.source.label} content={report.example.source.content} />
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
      setError(null);
    } catch (nextError) {
      setError(String(nextError));
    } finally {
      setLoading(false);
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
              <ReportView report={report} />
            </div>
          </ScrollArea>
        </div>
      </div>
    </main>
  );
}
