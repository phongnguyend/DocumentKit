import { useEffect, useRef, useState } from 'react';
import { Download, FileDown, Eye, LoaderCircle, X } from 'lucide-react';
import { Modal } from './Modal';
import './PdfExportDialog.css';

type Props = { bytes: Blob; name: string; page: number; count: number; onClose: () => void };
export function PdfExportDialog({ bytes, name, page, count, onClose }: Props) {
  const [selection, setSelection] = useState('all');
  const [includeHidden, setIncludeHidden] = useState(true);
  const [strict, setStrict] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [result, setResult] = useState<{ url: string; count: number; warnings: string[] } | null>(null);
  const [preview, setPreview] = useState(false);
  const controller = useRef<AbortController | null>(null);
  const objectUrl = useRef('');
  useEffect(() => () => {
    controller.current?.abort();
    URL.revokeObjectURL(objectUrl.current);
  }, []);
  function reset() {
    URL.revokeObjectURL(objectUrl.current);
    objectUrl.current = '';
    setResult(null);
    setPreview(false);
    setError('');
  }
  async function convert() {
    reset();
    setBusy(true);
    controller.current = new AbortController();
    try {
      const params = new URLSearchParams({ includeHidden: String(includeHidden), strict: String(strict) });
      if (selection === 'current') {
        params.set('slides', String(page));
      }
      const response = await fetch(`/api/documents/pdf?${params}`, { method: 'POST', body: bytes, signal: controller.current.signal, headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation' } });
      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.detail ?? 'PDF conversion failed.');
      }
      const binary = Uint8Array.from(atob(data.pdfBase64), character => character.charCodeAt(0));
      objectUrl.current = URL.createObjectURL(new Blob([binary], { type: 'application/pdf' }));
      setResult({ url: objectUrl.current, count: data.pageCount, warnings: data.warnings });
    } catch (failure) {
      if (!controller.current.signal.aborted) {
        setError(failure instanceof Error ? failure.message : 'PDF conversion failed.');
      }
    } finally {
      setBusy(false);
    }
  }
  return <Modal labelledBy="pdf-export-title" className={`pdf-export-modal${preview ? ' pdf-export-preview' : ''}`} onClose={onClose}>
    <header className="modal-heading"><div><h2 id="pdf-export-title"><FileDown size={22} />Export PDF</h2><p className="help">{name}</p></div><button aria-label="Close PDF export" onClick={onClose}><X size={18} /></button></header>
    <div className="pdf-export-body">
      <p className="help">Native conversion keeps supported text selectable and graphics sharp. Review the PDF and conversion notes before sharing.</p>
      <div className="pdf-export-options">
        <label>Slides<select disabled={busy} value={selection} onChange={event => { reset(); setSelection(event.target.value); }}><option value="all">All slides ({count})</option><option value="current">Current slide ({page})</option></select></label>
        <label className="pdf-export-check"><input type="checkbox" checked={includeHidden} disabled={busy} onChange={event => { reset(); setIncludeHidden(event.target.checked); }} /><span>Include hidden slides<small>Export slides marked as hidden in PowerPoint.</small></span></label>
        <label className="pdf-export-check"><input type="checkbox" checked={strict} disabled={busy} onChange={event => { reset(); setStrict(event.target.checked); }} /><span>Reject conversion warnings<small>Stop if a font is substituted or content cannot be fully rendered.</small></span></label>
      </div>
      {error && <p role="alert" className="error">{error}</p>}
      {busy && <p role="status">Converting slides...</p>}
      {result && <section className="pdf-export-result" aria-label="Conversion result"><strong>{result.count} PDF {result.count === 1 ? 'page' : 'pages'} ready</strong>{result.warnings.length > 0 ? <details><summary>Conversion notes ({result.warnings.length})</summary><ul>{result.warnings.map(warning => <li key={warning}>{warning}</li>)}</ul></details> : <p>No conversion warnings detected. Visual review is still recommended.</p>}</section>}
      {preview && result && <iframe title="Converted PDF preview" src={result.url} />}
    </div>
    <footer className="pdf-export-actions"><button onClick={onClose}>Close</button>{result ? <><button aria-pressed={preview} onClick={() => setPreview(!preview)}><Eye size={16} />{preview ? 'Hide preview' : 'Preview'}</button><a className="pdf-export-download" href={result.url} download={name.replace(/\.pptx$/i, '') + '.pdf'}><Download size={16} />Download PDF</a></> : <button className="primary" disabled={busy} onClick={() => void convert()}>{busy ? <LoaderCircle size={16} className="pdf-export-spinner" /> : <FileDown size={16} />}{busy ? 'Converting...' : 'Convert to PDF'}</button>}</footer>
  </Modal>;
}
