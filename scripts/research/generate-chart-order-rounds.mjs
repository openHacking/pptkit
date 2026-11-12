// Internal candidate research only: this does not bypass or authorize session saves.
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { DEFAULT_SESSION_LIMITS } from '../../packages/pptx-parser/dist/index.js';
import { structureContext } from '../../packages/pptx-parser/dist/presentation/structure-context.js';
import { readStructureFromContext } from '../../packages/pptx-parser/dist/presentation/read-structure.js';
import { projectStructure } from '../../packages/pptx-parser/dist/import/project-structure.js';
import { projectContent } from '../../packages/pptx-parser/dist/import/project-content.js';
import { snapshot } from '../../packages/pptx-parser/dist/import/state.js';
import { analyzeChartOrders } from '../../packages/pptx-parser/dist/import/analyze-chart-order.js';
import { analyzeTextEdits } from '../../packages/pptx-parser/dist/import/analyze-text-edits.js';
import { prepareTextCandidate } from '../../packages/pptx-parser/dist/import/prepare-text-candidate.js';

const directory = process.argv[2] ?? '.tmp/chart-order-research'; await mkdir(directory, { recursive: true });
const withLegend = process.argv[3] === 'legend';
let bytes = await readFile(withLegend ? 'tests/fixtures/chart-order-legend-before.pptx' : 'tests/fixtures/chart-label-source.pptx');
for (let round = 1; round <= 3; round++) {
  const context = await structureContext(bytes, undefined, true), inspection = await readStructureFromContext(context), checkpoint = async () => {};
  const projection = await projectStructure(inspection, DEFAULT_SESSION_LIMITS.maxReportFindings, checkpoint), labels = [], numbers = [], orders = [], args = Array(38).fill(undefined);
  [args[0], args[1], args[2], args[3], args[4], args[5], args[9], args[34], args[35], args[36], args[37]] = [projection.document, projection.report, context, DEFAULT_SESSION_LIMITS, checkpoint, {}, numbers, labels, true, orders, withLegend];
  await projectContent(...args);
  if (orders.length !== 1 || !orders[0].complete) throw Error('Chart order source is incomplete.');
  const baseline = snapshot(projection.document, DEFAULT_SESSION_LIMITS.maxSnapshotValues), candidate = structuredClone(baseline);
  const chart = candidate.slides[0].elements.find(item => item.type === 'chart');
  chart.series[0].name = `Order round ${round}`; chart.series[0].values[0] = withLegend ? round + 0.5 : 10 + round; chart.categories[1] = `Category ${round}`; chart.series.reverse();
  const binding = orders[0], identities = [...binding.sources].sort((a, b) => a.order - b.order).map(item => item.identity).reverse();
  const mapped = await analyzeChartOrders(baseline, candidate, orders, [{ slideId: binding.slideId, elementId: binding.elementId, identities }], DEFAULT_SESSION_LIMITS.maxSaveAnalysisWork, checkpoint);
  const scalar = await analyzeTextEdits(baseline, mapped.candidateForScalars, [], [], DEFAULT_SESSION_LIMITS, checkpoint, numbers, undefined, [], [], [], labels);
  if (scalar.diagnostics.length) throw Error(JSON.stringify(scalar.diagnostics));
  const result = await prepareTextCandidate(bytes, candidate, { ...scalar, chartOrders: mapped.proposals, changes: [...scalar.changes, ...mapped.changes] }, {}, DEFAULT_SESSION_LIMITS, checkpoint);
  if (!('bytes' in result)) throw Error(JSON.stringify(result));
  bytes = result.bytes; await writeFile(`${directory}/round-${round}.pptx`, bytes);
}
