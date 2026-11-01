import { useEffect, useRef, useState } from 'react';
import { ChevronLeft, ChevronRight } from 'lucide-react';
import { createPptxPreview } from '../renderers/pptxRenderer';
import './PptxPreview.css';

export function PptxPreview({ base64 }: { base64: string }) {
  const [deck, setDeck] = useState<Awaited<ReturnType<typeof createPptxPreview>> | null>(null);
  const [mode, setMode] = useState<'canvas' | 'selectable'>('canvas');
  const [page, setPage] = useState(1);
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(true);
  const host = useRef<HTMLDivElement>(null);
  useEffect(() => {
    let cancelled = false;
    setError('');
    setDeck(null);
    setPage(1);
    void createPptxPreview(base64).then(value => {
      if (!cancelled) {
        setDeck(value);
      }
    }).catch(cause => {
      if (!cancelled) {
        setError(cause instanceof Error ? cause.message : 'Could not open PowerPoint preview.');
        setBusy(false);
      }
    });
    return () => {
      cancelled = true;
    };
  }, [base64]);
  useEffect(() => {
    if (!deck) {
      return;
    }
    const controller = new AbortController();
    setBusy(true);
    setError('');
    const canvas = document.createElement('canvas');
    const scale = Math.min(2, 1600 / deck.width, 1600 / deck.height);
    canvas.width = Math.ceil(deck.width * scale);
    canvas.height = Math.ceil(deck.height * scale);
    canvas.setAttribute('aria-label', `PowerPoint slide ${page}`);
    canvas.setAttribute('role', 'img');
    const stage = document.createElement('div');
    stage.className = 'pptx-preview-stage';
    stage.append(canvas);
    const layer = document.createElement('div');
    layer.className = 'pptx-preview-text';
    layer.style.width = `${canvas.width}px`;
    layer.style.height = `${canvas.height}px`;
    layer.setAttribute('aria-label', `Selectable slide ${page} text`);
    const observer = new ResizeObserver(() => {
      layer.style.transform = `scale(${canvas.clientWidth / canvas.width}, ${canvas.clientHeight / canvas.height})`;
    });
    const fit = new ResizeObserver(() => {
      const container = host.current;
      if (container && window.matchMedia('(min-width: 900px)').matches) {
        const factor = Math.min(container.clientWidth / canvas.width, container.clientHeight / canvas.height);
        canvas.style.width = `${canvas.width * factor}px`;
        canvas.style.height = `${canvas.height * factor}px`;
      } else {
        canvas.style.removeProperty('width');
        canvas.style.removeProperty('height');
      }
    });
    void deck.render(canvas, page, controller.signal, mode === 'selectable' ? layer : undefined).then(() => {
      if (!controller.signal.aborted) {
        if (mode === 'selectable') {
          stage.append(layer);
          observer.observe(canvas);
        }
        host.current?.replaceChildren(stage);
        if (host.current) {
          fit.observe(host.current);
        }
      }
    }).catch(cause => {
      if (!controller.signal.aborted) {
        host.current?.replaceChildren();
        setError(cause instanceof Error ? cause.message : 'Could not render slide.');
      }
    }).finally(() => {
      if (!controller.signal.aborted) {
        setBusy(false);
      }
    });
    return () => {
      controller.abort();
      observer.disconnect();
      fit.disconnect();
    };
  }, [deck, page, mode]);

  return <section className="pptx-preview" aria-label="PowerPoint preview" aria-busy={busy}>
    <div className="pptx-preview-toolbar"><strong>Slide preview</strong><div>
      <button aria-label="Previous preview slide" disabled={!deck || page === 1 || busy} onClick={() => setPage(value => value - 1)}><ChevronLeft size={16} /></button>
      <span>Slide {page} of {deck?.count ?? '…'}</span>
      <button aria-label="Next preview slide" disabled={!deck || page === deck.count || busy} onClick={() => setPage(value => value + 1)}><ChevronRight size={16} /></button>
    </div></div>
    <div className="pptx-preview-modes" role="group" aria-label="Preview rendering mode">
      <button aria-pressed={mode === 'canvas'} disabled={busy} onClick={() => setMode('canvas')}>Canvas</button>
      <button aria-pressed={mode === 'selectable'} disabled={busy} onClick={() => setMode('selectable')}>Selectable text</button>
    </div>
    {busy && <p role="status">Rendering slide…</p>}
    {error && <p role="alert" className="assembly-error">{error}</p>}
    <div ref={host} className="pptx-preview-canvas" />
    <p className="control-help">Preview of the generated PPTX with your selected options. Browser fonts and text layout may differ from Microsoft PowerPoint.</p>
  </section>;
}
