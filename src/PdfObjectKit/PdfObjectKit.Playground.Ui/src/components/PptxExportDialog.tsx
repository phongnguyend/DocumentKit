import { useEffect, useRef, useState } from 'react';
import { Download, Presentation, LoaderCircle, X } from 'lucide-react';
import { Modal } from './Modal';
import './PageAssemblyDialog.css';
import './CompressionDialog.css';

type Result = { base64: string; slideCount: number; textBoxCount: number; imageCount: number; shapeCount: number; warnings: string[] };

export function PptxExportDialog({ pdf, name, pageCount, currentPage, onClose }: { pdf: Blob; name: string; pageCount: number; currentPage: number; onClose: () => void }) {
  const [scope, setScope] = useState('all');
  const [size, setSize] = useState('page');
  const [includeImages, setIncludeImages] = useState(true);
  const [includePaths, setIncludePaths] = useState(true);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [result, setResult] = useState<Result | null>(null);
  const abort = useRef<AbortController | null>(null);
  useEffect(() => () => abort.current?.abort(), []);

  async function convert() {
    const controller = new AbortController();
    abort.current = controller;
    setBusy(true);
    setError('');
    setResult(null);
    try {
      const options = { includeImages, includePaths, pageNumbers: scope === 'current' ? [currentPage] : null,
        ...(size === 'wide' ? { slideWidth: 960, slideHeight: 540 } : {}) };
      const response = await fetch(`/api/documents/pptx?options=${encodeURIComponent(JSON.stringify(options))}`, {
        method: 'POST', headers: { 'Content-Type': 'application/pdf' }, body: pdf, signal: controller.signal,
      });
      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.detail ?? 'Could not export PowerPoint.');
      }
      setResult(data);
    } catch (cause) {
      if (!controller.signal.aborted) {
        setError(cause instanceof Error ? cause.message : 'Could not export PowerPoint.');
      }
    } finally {
      setBusy(false);
    }
  }

  return <Modal labelledBy="pptx-title" className="page-assembly-dialog compression-dialog" onClose={onClose} canClose={!busy}>
    <div className="assembly-heading"><div><Presentation size={22} aria-hidden="true" /><h2 id="pptx-title">Export editable PowerPoint</h2></div><button aria-label="Close PowerPoint export" disabled={busy} onClick={onClose}><X size={18} /></button></div>
    <p className="assembly-document">{name} &middot; {pageCount} {pageCount === 1 ? 'page' : 'pages'}</p>
    <p>Convert PDF pages into slides with editable text lines, pictures, and supported vector shapes. Text on the same line is combined while keeping its formatting.</p>
    <fieldset className="compression-options" disabled={busy} onChange={() => setResult(null)}><legend>Slide content</legend>
      <label>Pages<select aria-label="Pages" value={scope} onChange={event => setScope(event.target.value)}><option value="all">All pages</option><option value="current">Current page ({currentPage})</option></select></label>
      <label>Slide size<select aria-label="Slide size" value={size} onChange={event => setSize(event.target.value)}><option value="page">Match first selected page</option><option value="wide">Widescreen (16:9)</option></select></label>
      <label><input type="checkbox" checked={includeImages} onChange={event => setIncludeImages(event.target.checked)} />Include pictures</label>
      <label><input type="checkbox" checked={includePaths} onChange={event => setIncludePaths(event.target.checked)} />Include vector shapes</label>
    </fieldset>
    <p className="control-help">Uses the current saved PDF. Mixed page sizes are fitted proportionally. Font substitution, clipping, and complex effects may change the appearance. Scanned pages remain pictures unless OCR text has been applied as visible replacement text.</p>
    {error && <p role="alert" className="assembly-error">{error}</p>}
    {result && <section className="compression-result" aria-label="PowerPoint export result" aria-live="polite">
      <strong>{result.slideCount} {result.slideCount === 1 ? 'slide' : 'slides'} ready</strong>
      <p>{result.textBoxCount} editable text boxes &middot; {result.imageCount} pictures &middot; {result.shapeCount} vector shapes</p>
      <details><summary>Conversion notes ({result.warnings.length})</summary>{result.warnings.map(warning => <p className="control-help" key={warning}>{warning}</p>)}</details>
    </section>}
    <div className="assembly-actions"><button onClick={() => {
      if (busy) {
        abort.current?.abort();
      }
      onClose();
    }}>{busy ? 'Cancel export' : 'Close'}</button>
      {!result && <button className="primary-action" disabled={busy} onClick={() => void convert()}>{busy ? <LoaderCircle size={16} /> : <Presentation size={16} />}{busy ? 'Converting...' : 'Convert to PowerPoint'}</button>}
      {result && <button className="primary-action" onClick={() => {
        const bytes = Uint8Array.from(atob(result.base64), character => character.charCodeAt(0));
        const url = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }));
        const link = document.createElement('a');
        link.href = url;
        link.download = name.replace(/\.pdf$/i, '') + '.pptx';
        link.click();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
      }}><Download size={16} />Download PPTX</button>}
    </div>
  </Modal>;
}
