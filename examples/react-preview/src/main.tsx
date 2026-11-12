import { useEffect, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { createPresentation, normalizePresentation, validatePresentation } from '@pptkit/core';
import { resolveNormalizedLayout } from '@pptkit/layout';
import { generatePptx } from '@pptkit/pptx-exporter';
import { renderIncrementalSvg, type IncrementalSvgCache } from '@pptkit/svg-renderer';

function createDeck(title: string) {
  const deck = createPresentation({ size: { width: 960, height: 540 } });
  deck.addSlide({ id: 'slide' }).addElement({ type: 'text', content: [{ runs: [{ text: title, style: { fontSize: 36, bold: true } }] }], box: { x: 60, y: 60, width: 840, height: 100 } });
  deck.addSlide({ id: 'details' }).addElement({ type: 'text', content: 'Supporting details', box: { x: 60, y: 60, width: 840, height: 100 } });
  return deck;
}

function App() {
  const [title, setTitle] = useState('Quarterly review');
  const [svg, setSvg] = useState<string[]>([]);
  const cache = useRef<IncrementalSvgCache | undefined>(undefined);
  const [status, setStatus] = useState('Preparing preview…');
  const [busy, setBusy] = useState(false);
  useEffect(() => {
    const controller = new AbortController();
    const deck = createDeck(title);
    const errors = validatePresentation(deck).filter(item => item.severity === 'error');
    if (errors.length) { setSvg([]); setStatus(JSON.stringify(errors)); return () => controller.abort(); }
    setStatus('Preparing preview…');
    let layout;
    try { layout = resolveNormalizedLayout(normalizePresentation(deck)); }
    catch (error) { setSvg([]); setStatus(String(error)); return () => controller.abort(); }
    void renderIncrementalSvg(layout, { signal: controller.signal, environmentRevision: 'react-example-v1' }, cache.current).then(({ result, cache: nextCache }) => {
      if (controller.signal.aborted) return;
      cache.current = nextCache;
      setSvg(result.slides.map(slide => slide.svg));
      setStatus(result.warnings.length ? JSON.stringify(result.warnings) : 'Preview ready');
    }).catch(error => { if (!controller.signal.aborted) { setSvg([]); setStatus(String(error)); } });
    return () => controller.abort();
  }, [title]);
  async function download() {
    setBusy(true);
    try {
      const result = await generatePptx(createDeck(title));
      const url = URL.createObjectURL(new Blob([new Uint8Array(result.bytes)], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }));
      const link = document.createElement('a'); link.href = url; link.download = 'review.pptx'; link.click();
      setTimeout(() => URL.revokeObjectURL(url), 0);
      setStatus(result.warnings.length ? JSON.stringify(result.warnings) : 'Download ready');
    } catch (error) { setStatus(String(error)); }
    finally { setBusy(false); }
  }
  return <main style={{ maxWidth: 960, margin: '2rem auto', fontFamily: 'sans-serif' }}>
    <h1>React presentation preview</h1>
    <label>Slide title <input value={title} onChange={event => setTitle(event.target.value)} /></label>{' '}
    <button type="button" onClick={() => void download()} disabled={busy}>Download PPTX</button>
    <p role="status">{status}</p>
    {svg.map((slide, index) => <iframe key={index} title={index === 0 ? 'Slide preview' : `Slide preview ${index + 1}`} sandbox="" srcDoc={slide} style={{ width: '100%', height: 560, border: '1px solid #ccc' }} />)}
  </main>;
}

const root = document.getElementById('root');
if (!root) throw new Error('Missing React root.');
createRoot(root).render(<App />);
