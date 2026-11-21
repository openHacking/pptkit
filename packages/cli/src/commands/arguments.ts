import { DEFAULT_ZIP_LIMITS, DEFAULT_NOTES_INSPECTION_LIMITS } from "@pptkit/pptx-parser";
import type { OpcComparisonProfile, PptxImportOptions, PptxNotesInspectionLimits } from "@pptkit/pptx-parser";
export const HELP = `PPTKit CLI
Usage:
  pptkit inspect <file> [--projection structure|content] [--json]
  pptkit compare <before> <after> [--profile opc-payload-v1|opc-xml-v1] [--json]
  pptkit save <source> <destination> --mode no-op [--json]
  pptkit save <source> <destination> --mode edit --first-slide-number N|clear [--json]
  pptkit save <source> <destination> --mode edit --text-edits edits.json [--json]
  pptkit save <source> <destination> --mode edit --chart-edits charts.json [--json]
  pptkit save <source> <destination> --mode edit --description-edits descriptions.json [--json]
  pptkit save <source> <destination> --mode edit --name-edits names.json [--json]
  pptkit save <source> <destination> --mode edit --decorative-edits decorative.json [--json]
  pptkit save <source> <destination> --mode edit --visibility-edits visibility.json [--json]
  pptkit save <source> <destination> --mode edit --metadata-edits metadata.json [--json]
  pptkit save <source> <destination> --mode edit --custom-property-edits properties.json [--json]
  pptkit save <source> <destination> --mode edit --picture-edit picture.json --picture-bytes image [--json]
  pptkit generate <authoring.json> <destination> [--fail-on-export-warning] [--json]
  pptkit validate <authoring.json> [--export [--fail-on-export-warning]] [--json]
  pptkit preview <file> --slide N|all [--input-format authoring-json|pptx] [--output FILE] --json
Options:
  --max-input-bytes N        Per file; default 67108864, maximum 536870912
  --asset-base DIR           Explicit base directory for local authoring assets
  --output FILE              Create a single SVG or all-slides HTML preview
  --max-output-bytes N       UTF-8 output; default 1048576, range 512..67108864
  --max-differences N        Compare findings retained; default 1000
  --diagram-preview cache   Inspect content projection only; freshness unknown
  --diagram-text-comparison direct-paragraphs
  --edit-save-profile text-runs-v1|text-runs-v2|text-runs-v3
  --chart-edit-save-profile existing-values-v1|existing-values-v2|existing-values-v3|existing-values-and-labels-v4|existing-values-and-series-names-v5
  --picture-source-profile embedded-v1
  --capabilities            Inspect a seven-operation capability view
  --package-summary         Add declared-media sizes and relationship counts (inspect JSON only)
  --notes-inventory         Add source-only notes-page inventory (inspect only)
  --max-notes-work N        Notes inventory work units; default 1000000
  --max-notes-records N     Notes inventory records; default 100000
  --max-notes-findings N    Notes inventory findings; default 10000
  --max-notes-text N        Notes inventory UTF-16 code units; default 1000000
                            Notes limits require --notes-inventory; range 1..100000000
  --export                  Validate export and generated package in memory (validate only)
  --check-descriptions      Inspect missing image/chart/table descriptions (validate only)
  --description-occurrences  Report modeled slide occurrences (requires check)
  --fail-on-missing-description  Make missing descriptions fail validation (requires check)
  --check-contrast          Inspect text contrast and indeterminate cases (validate only)
  --fail-on-contrast        Fail on below-threshold or indeterminate text (requires check)
  --text-values             Include bound text-edit targets and baseline text (inspect JSON only)
  --chart-values            Include bound chart-edit targets and numeric baselines (inspect JSON only)
  --description-values      Include bound object descriptions (inspect JSON only)
  --name-values             Include bound object names (inspect JSON only)
  --decorative-values       Include bound decorative values (inspect JSON only)
  --visibility-values       Include bound slide visibility (inspect JSON only)
  --metadata-values         Include bound core metadata strings (inspect JSON only)
  --metadata-revision       Request v2 revision candidates (requires metadata-values)
  --metadata-keywords       Request v3 aggregate keyword candidates (requires metadata-values)
  --metadata-company        Request v4 company candidates (requires metadata-values)
  --metadata-insert         Request v5 missing-field candidates (requires metadata-values)
  --metadata-dates          Request v6 date candidates (requires metadata-values)
  --metadata-date-insert    Request v7 date insertions (requires metadata-values)
  --metadata-create         Request v8 core-part creation (requires metadata-values)
  --metadata-revision-insert  Request v9 missing revision (requires metadata-values)
  --metadata-keywords-insert  Request v10 missing keywords (requires metadata-values)
  --metadata-company-insert  Request v11 missing company (requires metadata-values)
  --metadata-extended-create Request v12 extended part creation (requires metadata-values)
  --metadata-remove          Request v13 metadata removal (requires metadata-values)
  --metadata-date-remove     Request v14 date removal (requires metadata-values)
  --metadata-keywords-remove Request v15 keyword removal (requires metadata-values)
  --metadata-revision-remove Request v16 revision removal (requires metadata-values)
  --metadata-default-remove  Request v17 defaulted metadata removal (requires metadata-values)
  --custom-property-values  Include bound custom scalar values (inspect JSON only)
  --custom-property-currency Request exact currency values (v6; v8 with decimal) (requires custom-property-values)
  --custom-property-decimal Request exact decimal edits (v7; v8 with currency) (requires custom-property-values)
  --custom-property-names   Request v2 rename candidates (requires custom-property-values)
  --custom-property-deletions  Request v3 deletion candidates (requires custom-property-values)
  --custom-property-additions  Request v4 insertion capability (requires custom-property-values)
  --custom-property-create  Request v5 creation capability (requires custom-property-additions)
  --dry-run                 Prepare and report a save without publishing (save only)
  --fail-on-export-warning  Reject export warnings before publishing
  --image-diagnostics       Check PNG/GIF structure and JPEG metadata on export
  --validate-image-data     Omit images that fail PNG, GIF or JPEG data checks
  --max-capability-subjects N  Complete subjects retained; default 1000, zero allowed
  --capability-feature NAME  Exact feature filter (requires --capabilities)
  --capability-part PART     Exact logical part filter (requires --capabilities)
  --help                    Show help; -- ends option parsing
Exit: 0 success/equal, 1 different or validation errors, 2 usage, 3 invalid/unsupported input or I/O,
      4 resource limit, 5 internal error, 130 cancelled.
Save creates a new destination only. Local assets require --asset-base; no network loading or native verification.`;
export class CliError extends Error {
  constructor(readonly code: string, message: string, readonly exitCode: number, readonly source?: unknown) { super(message); }
}
export interface Command {
  kind: "inspect" | "compare" | "save" | "generate" | "validate" | "preview";
  slideNumber?: number | "all";
  previewInput?: "authoring-json" | "pptx";
  saveMode?: "no-op" | "edit";
  firstSlideNumber?: number | null;
  textEdits?: string;
  chartEdits?: string;
  descriptionEdits?: string;
  nameEdits?: string;
  decorativeEdits?: string;
  visibilityEdits?: string;
  metadataEdits?: string;
  customPropertyEdits?: string;
  pictureEdit?: string;
  pictureBytes?: string;
  files: string[];
  json: boolean;
  validateExport: boolean;
  failOnExportWarning: boolean;
  imageDiagnostics: boolean;
  validateImageData: boolean;
  checkDescriptions: boolean;
  descriptionOccurrences: boolean;
  failOnMissingDescription: boolean;
  checkContrast: boolean;
  failOnContrast: boolean;
  textValues: boolean;
  chartValues: boolean;
  descriptionValues: boolean;
  nameValues: boolean;
  decorativeValues: boolean;
  visibilityValues: boolean;
  metadataValues: boolean;
  metadataRevision: boolean;
  metadataKeywords: boolean;
  metadataCompany: boolean;
  metadataInsert: boolean;
  metadataDates: boolean;
  metadataDateInsert: boolean;
  metadataCreate: boolean;
  metadataRevisionInsert: boolean;
  metadataKeywordsInsert: boolean;
  metadataCompanyInsert: boolean;
  metadataExtendedCreate: boolean;
  metadataRemove: boolean;
  metadataDateRemove: boolean;
  metadataKeywordsRemove: boolean;
  metadataRevisionRemove: boolean;
  metadataDefaultRemove: boolean;
  customPropertyValues: boolean;
  customPropertyCurrency: boolean;
  customPropertyDecimal: boolean;
  customPropertyNames: boolean;
  customPropertyDeletions: boolean;
  customPropertyAdditions: boolean;
  customPropertyCreate: boolean;
  dryRun: boolean;
  assetBase?: string;
  previewOutput?: string;
  maxInputBytes: number;
  maxOutputBytes: number;
  projection: "structure" | "content";
  profile: OpcComparisonProfile;
  maxDifferences: number;
  capabilities: boolean;
  packageSummary: boolean;
  notesInventory: boolean;
  notesLimits: PptxNotesInspectionLimits;
  maxCapabilitySubjects: number;
  capabilityFeature?: string;
  capabilityPart?: string;
  diagramPreview?: "cache";
  diagramTextComparison?: "direct-paragraphs";
  editSaveProfile?: PptxImportOptions["editSaveProfile"];
  chartEditSaveProfile?: PptxImportOptions["chartEditSaveProfile"];
  pictureSourceProfile?: PptxImportOptions["pictureSourceProfile"];
}
const usage = (message: string): never => { throw new CliError("usage", message, 2); };
export function parseCommand(args: readonly string[]): Command | undefined {
  if (!Array.isArray(args) || args.some(a => typeof a !== "string" || !a.length || a.length > 32768 || a.includes("\0"))) return usage("Arguments must be nonempty bounded strings without NUL.");
  if (!args.length || args.length === 1 && args[0] === "--help" || args.length === 2 && ["inspect", "compare", "save", "generate", "validate", "preview"].includes(args[0]!) && args[1] === "--help") return undefined;
  const kind = args[0]; if (kind !== "inspect" && kind !== "compare" && kind !== "save" && kind !== "generate" && kind !== "validate" && kind !== "preview") return usage("Expected inspect, compare, save, generate, validate or preview. Use --help.");
  const options = new Map<string, string>(), files: string[] = []; let positional = false;
  const flags = new Set(["--input-format", "--slide", "--mode", "--first-slide-number", "--text-edits", "--chart-edits", "--description-edits", "--name-edits", "--decorative-edits", "--visibility-edits", "--metadata-edits", "--custom-property-edits", "--picture-edit", "--picture-bytes", "--picture-source-profile", "--chart-edit-save-profile", "--projection", "--profile", "--max-input-bytes", "--max-output-bytes", "--max-differences", "--max-notes-work", "--max-notes-records", "--max-notes-findings", "--max-notes-text", "--diagram-preview", "--diagram-text-comparison", "--edit-save-profile", "--max-capability-subjects", "--capability-feature", "--capability-part", "--asset-base", "--output"]);
  for (let i = 1; i < args.length; i++) {
    const value = args[i]!;
    if (!positional && value === "--") { positional = true; continue; }
    if (!positional && value.startsWith("-")) {
      if (options.has(value)) return usage("Duplicate option.");
      if (["--json", "--capabilities", "--package-summary", "--notes-inventory", "--export", "--fail-on-export-warning", "--image-diagnostics", "--validate-image-data", "--check-descriptions", "--description-occurrences", "--fail-on-missing-description", "--check-contrast", "--fail-on-contrast", "--text-values", "--chart-values", "--description-values", "--name-values", "--decorative-values", "--visibility-values", "--metadata-values", "--metadata-revision", "--metadata-keywords", "--metadata-company", "--metadata-insert", "--metadata-dates", "--metadata-date-insert", "--metadata-create", "--metadata-revision-insert", "--metadata-keywords-insert", "--metadata-company-insert", "--metadata-extended-create", "--metadata-remove", "--metadata-date-remove", "--metadata-keywords-remove", "--metadata-revision-remove", "--metadata-default-remove", "--custom-property-values", "--custom-property-currency", "--custom-property-decimal", "--custom-property-names", "--custom-property-deletions", "--custom-property-additions", "--custom-property-create", "--dry-run"].includes(value)) { options.set(value, "true"); continue; }
      if (!flags.has(value)) return usage("Unknown option. Use --help.");
      const next = args[++i]; if (!next || next.startsWith("--")) return usage("Option requires a value.");
      options.set(value, next);
    } else files.push(value);
  }
  if (files.length !== (kind === "inspect" || kind === "validate" || kind === "preview" ? 1 : 2)) return usage("Incorrect number of input files.");
  const only = kind === "inspect" ? ["--mode", "--first-slide-number", "--profile", "--max-differences"] : kind === "compare" ? ["--mode", "--first-slide-number", "--projection", "--diagram-preview", "--diagram-text-comparison", "--edit-save-profile", "--capabilities", "--max-capability-subjects", "--capability-feature", "--capability-part"] : ["--projection", "--profile", "--max-differences", "--diagram-preview", "--diagram-text-comparison", "--edit-save-profile", "--capabilities", "--max-capability-subjects", "--capability-feature", "--capability-part", ...(kind === "generate" || kind === "validate" || kind === "preview" ? ["--mode", "--first-slide-number"] : [])];
  if (only.some(k => options.has(k))) return usage("Option is not valid for this command.");
  if (options.has("--package-summary") && (kind !== "inspect" || !options.has("--json"))) return usage("--package-summary requires inspect --json.");
  if (options.has("--notes-inventory") && kind !== "inspect") return usage("--notes-inventory requires inspect.");
  if (!options.has("--notes-inventory") && ["--max-notes-work", "--max-notes-records", "--max-notes-findings", "--max-notes-text"].some(key => options.has(key))) return usage("Notes limits require --notes-inventory.");
  if (kind !== "validate" && options.has("--export")) return usage("--export is validate-only.");
  if (kind !== "validate" && (options.has("--check-descriptions") || options.has("--description-occurrences") || options.has("--fail-on-missing-description")) || (options.has("--description-occurrences") || options.has("--fail-on-missing-description")) && !options.has("--check-descriptions")) return usage("Description checking is validate-only; occurrence and failure options require --check-descriptions.");
  if (kind !== "validate" && (options.has("--check-contrast") || options.has("--fail-on-contrast")) || options.has("--fail-on-contrast") && !options.has("--check-contrast")) return usage("Contrast checking is validate-only; failure policy requires --check-contrast.");
  if (options.has("--fail-on-export-warning") && kind !== "generate" && !(kind === "validate" && options.has("--export"))) return usage("--fail-on-export-warning requires generate or validate --export.");
  if (options.has("--validate-image-data") && kind !== "generate" && !(kind === "validate" && options.has("--export"))) return usage("--validate-image-data requires generate or validate --export.");
  if (options.has("--image-diagnostics") && kind !== "generate" && !(kind === "validate" && options.has("--export"))) return usage("--image-diagnostics requires generate or validate --export.");
  if (kind === "preview" ? !options.has("--slide") || !options.has("--json") : options.has("--slide")) return usage("Preview requires --slide N|all and --json; --slide is preview-only.");
  if (options.has("--output") && kind !== "preview") return usage("--output requires preview.");
  const previewInput = options.get("--input-format") ?? "authoring-json";
  if (kind !== "preview" && options.has("--input-format") || previewInput !== "authoring-json" && previewInput !== "pptx") return usage("Preview input format must be authoring-json or pptx.");
  if (options.has("--asset-base") && !(kind === "generate" || kind === "validate" || kind === "preview" && previewInput === "authoring-json")) return usage("--asset-base requires authoring JSON input.");
  const saveMode = options.get("--mode"), number = options.get("--first-slide-number");
  const textEdits = options.get("--text-edits"), chartEdits = options.get("--chart-edits"), descriptionEdits = options.get("--description-edits"), nameEdits = options.get("--name-edits"), decorativeEdits = options.get("--decorative-edits"), visibilityEdits = options.get("--visibility-edits"), metadataEdits = options.get("--metadata-edits"), customPropertyEdits = options.get("--custom-property-edits"), pictureEdit = options.get("--picture-edit"), pictureBytes = options.get("--picture-bytes"), pictureProfile = options.get("--picture-source-profile"), chartProfile = options.get("--chart-edit-save-profile");
  if (kind !== "save" && (textEdits !== undefined || chartEdits !== undefined || descriptionEdits !== undefined || nameEdits !== undefined || decorativeEdits !== undefined || visibilityEdits !== undefined || metadataEdits !== undefined || customPropertyEdits !== undefined || pictureEdit !== undefined || pictureBytes !== undefined)) return usage("Edit manifests are save-only.");
  if ((pictureEdit === undefined) !== (pictureBytes === undefined) || pictureEdit !== undefined && (textEdits !== undefined || chartEdits !== undefined || descriptionEdits !== undefined || nameEdits !== undefined || decorativeEdits !== undefined || visibilityEdits !== undefined || metadataEdits !== undefined || customPropertyEdits !== undefined)) return usage("Picture edits require both picture options and cannot mix with text, chart, description, name, decorative, visibility or metadata edits.");
  if (pictureProfile !== undefined && (kind !== "inspect" || options.get("--projection") !== "content" || pictureProfile !== "embedded-v1")) return usage("Picture source inspection requires content projection and embedded-v1.");
  if (kind !== "save" && options.has("--dry-run")) return usage("--dry-run is save-only.");
  if (chartProfile !== undefined && (kind !== "inspect" || options.get("--projection") !== "content" || !["existing-values-v1", "existing-values-v2", "existing-values-v3", "existing-values-and-labels-v4", "existing-values-and-series-names-v5"].includes(chartProfile))) return usage("Chart save inspection requires content projection and a supported chart profile.");
  if (kind === "save" && (saveMode !== "no-op" && saveMode !== "edit" || saveMode === "no-op" && (number !== undefined || textEdits !== undefined || chartEdits !== undefined || descriptionEdits !== undefined || nameEdits !== undefined || decorativeEdits !== undefined || visibilityEdits !== undefined || metadataEdits !== undefined || customPropertyEdits !== undefined || pictureEdit !== undefined) || saveMode === "edit" && number === undefined && textEdits === undefined && chartEdits === undefined && descriptionEdits === undefined && nameEdits === undefined && decorativeEdits === undefined && visibilityEdits === undefined && metadataEdits === undefined && customPropertyEdits === undefined && pictureEdit === undefined)) return usage("Save requires no-op or edit with at least one edit option.");
  if (number !== undefined && number !== "clear" && (!/^(?:0|[1-9][0-9]{0,3})$/.test(number))) return usage("First slide number must be 0..9999 or clear.");
  if (!options.has("--capabilities") && ["--max-capability-subjects", "--capability-feature", "--capability-part"].some(k => options.has(k))) return usage("Capability filters/limits require --capabilities.");
  const integer = (key: string, fallback: number, min: number, max: number): number => {
    const text = options.get(key); if (text === undefined) return fallback;
    const n = Number(text); if (!/^(?:0|[1-9][0-9]*)$/.test(text) || !Number.isSafeInteger(n) || n < min || n > max) return usage("Invalid numeric option."); return n;
  };
  const projection = options.get("--projection") ?? "structure";
  if (projection !== "structure" && projection !== "content") return usage("Unknown projection.");
  const profile = options.get("--profile") ?? "opc-payload-v1";
  if (profile !== "opc-payload-v1" && profile !== "opc-xml-v1") return usage("Unknown comparison profile.");
  const preview = options.get("--diagram-preview"), text = options.get("--diagram-text-comparison"), edit = options.get("--edit-save-profile");
  if (preview !== undefined && (preview !== "cache" || projection !== "content")) return usage("Cached preview requires content projection and cache.");
  if (text !== undefined && text !== "direct-paragraphs") return usage("Unknown diagram text comparison profile.");
  if (edit !== undefined && (projection !== "content" || !["text-runs-v1", "text-runs-v2", "text-runs-v3"].includes(edit))) return usage("Text save inspection requires content projection and a supported profile.");
  if (options.has("--text-values") && (kind !== "inspect" || !options.has("--json") || !edit || options.has("--capabilities"))) return usage("--text-values requires inspect --json with a text edit profile and no capability view.");
  if (options.has("--chart-values") && (kind !== "inspect" || !options.has("--json") || !chartProfile || options.has("--capabilities"))) return usage("--chart-values requires inspect --json with a chart edit profile and no capability view.");
  if (options.has("--description-values") && (kind !== "inspect" || !options.has("--json") || projection !== "content" || options.has("--capabilities") || edit !== undefined && edit !== "text-runs-v3" || pictureProfile !== undefined || preview !== undefined || text !== undefined)) return usage("--description-values requires JSON content inspection without incompatible profiles; text inspection must use text-runs-v3.");
  if (options.has("--name-values") && (kind !== "inspect" || !options.has("--json") || projection !== "content" || options.has("--capabilities") || edit !== undefined && edit !== "text-runs-v3" || pictureProfile !== undefined || preview !== undefined || text !== undefined)) return usage("--name-values requires JSON content inspection without incompatible profiles; text inspection must use text-runs-v3.");
  if (options.has("--decorative-values") && (kind !== "inspect" || !options.has("--json") || projection !== "content" || options.has("--capabilities") || edit !== undefined && edit !== "text-runs-v3" || pictureProfile !== undefined || preview !== undefined || text !== undefined)) return usage("--decorative-values requires JSON content inspection without incompatible profiles; text inspection must use text-runs-v3.");
  if (options.has("--visibility-values") && (kind !== "inspect" || !options.has("--json") || projection !== "content" || options.has("--capabilities") || pictureProfile !== undefined || preview !== undefined || text !== undefined)) return usage("--visibility-values requires JSON content inspection without incompatible profiles.");
  if (options.has("--metadata-values") && (kind !== "inspect" || !options.has("--json") || projection !== "content" || options.has("--capabilities") || pictureProfile !== undefined || preview !== undefined || text !== undefined)) return usage("--metadata-values requires JSON content inspection without incompatible profiles.");
  if (options.has("--metadata-default-remove") && !options.has("--metadata-values")) return usage("--metadata-default-remove requires --metadata-values.");
  if (options.has("--metadata-revision-remove") && !options.has("--metadata-values")) return usage("--metadata-revision-remove requires --metadata-values.");
  if (options.has("--metadata-keywords-remove") && !options.has("--metadata-values")) return usage("--metadata-keywords-remove requires --metadata-values.");
  if (options.has("--metadata-date-remove") && !options.has("--metadata-values")) return usage("--metadata-date-remove requires --metadata-values.");
  if (options.has("--metadata-remove") && !options.has("--metadata-values")) return usage("--metadata-remove requires --metadata-values.");
  if (options.has("--metadata-extended-create") && !options.has("--metadata-values")) return usage("--metadata-extended-create requires --metadata-values.");
  if (options.has("--metadata-company-insert") && !options.has("--metadata-values")) return usage("--metadata-company-insert requires --metadata-values.");
  if (options.has("--metadata-keywords-insert") && !options.has("--metadata-values")) return usage("--metadata-keywords-insert requires --metadata-values.");
  if (options.has("--metadata-revision-insert") && !options.has("--metadata-values")) return usage("--metadata-revision-insert requires --metadata-values.");
  if (options.has("--metadata-create") && !options.has("--metadata-values")) return usage("--metadata-create requires --metadata-values.");
  if (options.has("--metadata-date-insert") && !options.has("--metadata-values")) return usage("--metadata-date-insert requires --metadata-values.");
  if (options.has("--metadata-dates") && !options.has("--metadata-values")) return usage("--metadata-dates requires --metadata-values.");
  if (options.has("--metadata-insert") && !options.has("--metadata-values")) return usage("--metadata-insert requires --metadata-values.");
  if (options.has("--metadata-company") && !options.has("--metadata-values")) return usage("--metadata-company requires --metadata-values.");
  if (options.has("--metadata-keywords") && !options.has("--metadata-values")) return usage("--metadata-keywords requires --metadata-values.");
  if (options.has("--metadata-revision") && !options.has("--metadata-values")) return usage("--metadata-revision requires --metadata-values.");
  if (options.has("--custom-property-values") && (kind !== "inspect" || !options.has("--json") || projection !== "content" || options.has("--capabilities") || pictureProfile !== undefined || preview !== undefined || text !== undefined)) return usage("--custom-property-values requires JSON content inspection without incompatible profiles.");
  if ((options.has("--custom-property-names") || options.has("--custom-property-deletions") || options.has("--custom-property-additions")) && !options.has("--custom-property-values")) return usage("Custom property names/deletions require --custom-property-values.");
  if (options.has("--custom-property-currency") && (!options.has("--custom-property-values") || chartProfile !== undefined || ["--custom-property-names", "--custom-property-deletions", "--custom-property-additions", "--custom-property-create", "--chart-values", "--description-values", "--name-values", "--decorative-values"].some(option => options.has(option)))) return usage("--custom-property-currency requires --custom-property-values and compatible text, metadata or visibility profiles.");
  if (options.has("--custom-property-decimal") && (!options.has("--custom-property-values") || chartProfile !== undefined || ["--custom-property-names", "--custom-property-deletions", "--custom-property-additions", "--custom-property-create", "--chart-values", "--description-values", "--name-values", "--decorative-values"].some(option => options.has(option)))) return usage("--custom-property-decimal requires --custom-property-values and compatible text, metadata or visibility profiles.");
  if (options.has("--custom-property-create") && !options.has("--custom-property-additions")) return usage("--custom-property-create requires --custom-property-additions.");
  return { kind, files, json: options.has("--json"), validateExport: options.has("--export"), failOnExportWarning: options.has("--fail-on-export-warning"), imageDiagnostics: options.has("--image-diagnostics"), validateImageData: options.has("--validate-image-data"), checkDescriptions: options.has("--check-descriptions"), descriptionOccurrences: options.has("--description-occurrences"), failOnMissingDescription: options.has("--fail-on-missing-description"), checkContrast: options.has("--check-contrast"), failOnContrast: options.has("--fail-on-contrast"), textValues: options.has("--text-values"), chartValues: options.has("--chart-values"), descriptionValues: options.has("--description-values"), nameValues: options.has("--name-values"), decorativeValues: options.has("--decorative-values"), visibilityValues: options.has("--visibility-values"), metadataValues: options.has("--metadata-values"), metadataRevision: options.has("--metadata-revision"), metadataKeywords: options.has("--metadata-keywords"), metadataCompany: options.has("--metadata-company"), metadataInsert: options.has("--metadata-insert"), metadataDates: options.has("--metadata-dates"), metadataDateInsert: options.has("--metadata-date-insert"), metadataCreate: options.has("--metadata-create"), metadataRevisionInsert: options.has("--metadata-revision-insert"), metadataKeywordsInsert: options.has("--metadata-keywords-insert"), metadataCompanyInsert: options.has("--metadata-company-insert"), metadataExtendedCreate: options.has("--metadata-extended-create"), metadataRemove: options.has("--metadata-remove"), metadataDateRemove: options.has("--metadata-date-remove"), metadataKeywordsRemove: options.has("--metadata-keywords-remove"), metadataRevisionRemove: options.has("--metadata-revision-remove"), metadataDefaultRemove: options.has("--metadata-default-remove"), customPropertyValues: options.has("--custom-property-values"), customPropertyCurrency: options.has("--custom-property-currency"), customPropertyDecimal: options.has("--custom-property-decimal"), customPropertyNames: options.has("--custom-property-names"), customPropertyDeletions: options.has("--custom-property-deletions"), customPropertyAdditions: options.has("--custom-property-additions"), customPropertyCreate: options.has("--custom-property-create"), dryRun: options.has("--dry-run"), projection, profile,
    ...(options.has("--output") ? { previewOutput: options.get("--output")! } : {}),
    ...(options.has("--asset-base") ? { assetBase: options.get("--asset-base")! } : {}),
    ...(kind === "preview" ? { slideNumber: options.get("--slide") === "all" ? "all" as const : integer("--slide", 1, 1, 1000000), previewInput } : {}),
    ...(kind === "save" ? { saveMode: saveMode as "no-op" | "edit", ...(number !== undefined ? { firstSlideNumber: number === "clear" ? null : Number(number) } : {}), ...(textEdits !== undefined ? { textEdits } : {}), ...(chartEdits !== undefined ? { chartEdits } : {}), ...(descriptionEdits !== undefined ? { descriptionEdits } : {}), ...(nameEdits !== undefined ? { nameEdits } : {}), ...(decorativeEdits !== undefined ? { decorativeEdits } : {}), ...(visibilityEdits !== undefined ? { visibilityEdits } : {}), ...(metadataEdits !== undefined ? { metadataEdits } : {}), ...(customPropertyEdits !== undefined ? { customPropertyEdits } : {}), ...(pictureEdit !== undefined ? { pictureEdit, pictureBytes: pictureBytes! } : {}) } : {}),
    capabilities: options.has("--capabilities"), packageSummary: options.has("--package-summary"), notesInventory: options.has("--notes-inventory"),
    notesLimits: { maxWorkUnits: integer("--max-notes-work", DEFAULT_NOTES_INSPECTION_LIMITS.maxWorkUnits, 1, 100000000), maxRecords: integer("--max-notes-records", DEFAULT_NOTES_INSPECTION_LIMITS.maxRecords, 1, 100000000), maxFindings: integer("--max-notes-findings", DEFAULT_NOTES_INSPECTION_LIMITS.maxFindings, 1, 100000000), maxTextCharacters: integer("--max-notes-text", DEFAULT_NOTES_INSPECTION_LIMITS.maxTextCharacters, 1, 100000000) },
    maxCapabilitySubjects: integer("--max-capability-subjects", 1000, 0, 150000),
    ...(options.has("--capability-feature") ? { capabilityFeature: options.get("--capability-feature")! } : {}), ...(options.has("--capability-part") ? { capabilityPart: options.get("--capability-part")! } : {}),
    maxInputBytes: integer("--max-input-bytes", DEFAULT_ZIP_LIMITS.maxInputBytes, 1, 512 * 1024 * 1024),
    maxOutputBytes: integer("--max-output-bytes", 1024 * 1024, 512, 64 * 1024 * 1024), maxDifferences: integer("--max-differences", 1000, 1, 1000000),
    ...(preview ? { diagramPreview: "cache" } : {}), ...(text ? { diagramTextComparison: "direct-paragraphs" } : {}),
    ...(edit ? { editSaveProfile: edit as Command["editSaveProfile"] } : {}), ...(chartProfile ? { chartEditSaveProfile: chartProfile as Command["chartEditSaveProfile"] } : {}), ...(pictureProfile ? { pictureSourceProfile: "embedded-v1" } : {}) };
}
