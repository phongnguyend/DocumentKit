import { useEffect, useRef, useState } from 'react';
import { Download, FileText, LoaderCircle, X, Image, Files } from 'lucide-react';
import { Modal } from './Modal';
import './PageAssemblyDialog.css';
import './CompressionDialog.css';
import './PptxExportDialog.css';

type Result = { base64: string; sourcePageCount: number; paragraphCount: number; tableCount: number; imageCount: number; warnings: string[] };

export function DocxExportDialog({ pdf, name, pageCount, currentPage, onClose }: { pdf: Blob; name: string; pageCount: number; currentPage: number; onClose: () => void }) {
  const [scope, setScope] = useState('all');
  const [includeImages, setIncludeImages] = useState(true);
  const [pageBreaks, setPageBreaks] = useState(true);
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
      const options = { includeImages, pageBreaks, pageNumbers: scope === 'current' ? [currentPage] : null };
      const response = await fetch(`/api/documents/docx?options=${encodeURIComponent(JSON.stringify(options))}`, {
        method: 'POST', headers: { 'Content-Type': 'application/pdf' }, body: pdf, signal: controller.signal,
      });
      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.detail ?? 'Could not export Word.');
      }
      setResult(data);
    } catch (cause) {
      if (!controller.signal.aborted) {
        setError(cause instanceof Error ? cause.message : 'Could not export Word.');
      }
    } finally {
      if (!controller.signal.aborted) {
        setBusy(false);
      }
    }
  }

  return <Modal labelledBy="docx-title" className="page-assembly-dialog compression-dialog pptx-export-dialog" onClose={onClose} canClose={!busy}>
    <div className="assembly-heading"><div><FileText size={22} aria-hidden="true" /><h2 id="docx-title">Export editable Word</h2></div><button aria-label="Close Word export" disabled={busy} onClick={onClose}><X size={18} /></button></div>
    <div className="pptx-export-body"><div className="pptx-export-settings">
      <p className="assembly-document">{name} &middot; {pageCount} pages</p>
      <p className="pptx-export-intro">Create editable paragraphs, headings, tables and pictures. Word reflows the content, so layout and page counts may change.</p>
      <fieldset className="pptx-options" disabled={busy} onChange={() => setResult(null)}><legend>Word content</legend>
        <div className="pptx-option-fields"><label><span>Pages</span><select aria-label="Pages" value={scope} onChange={event => setScope(event.target.value)}><option value="all">All pages ({pageCount})</option><option value="current">Current page ({currentPage})</option></select><small>Choose up to 100 source pages.</small></label></div>
        <div className="pptx-content-choices">
          <label className="pptx-content-choice"><Image size={20} aria-hidden="true" /><span><strong>Include pictures</strong><small>Full source images; PDF cropping is not retained.</small></span><input type="checkbox" checked={includeImages} onChange={event => setIncludeImages(event.target.checked)} aria-label="Include pictures" /></label>
          <label className="pptx-content-choice"><Files size={20} aria-hidden="true" /><span><strong>Keep source page breaks</strong><small>Start each PDF page on a new Word page.</small></span><input type="checkbox" checked={pageBreaks} onChange={event => setPageBreaks(event.target.checked)} aria-label="Keep source page breaks" /></label>
        </div>
      </fieldset>
      <p className="control-help">Uses the current saved PDF. Scanned images require OCR for editable text. Vector artwork and exact visual formatting are not preserved. Review conversion notes before sharing.</p>
      {error && <p role="alert" className="assembly-error">{error}</p>}
      {result && <section className="compression-result" aria-label="Word export result" aria-live="polite">
        <strong>{result.sourcePageCount} source {result.sourcePageCount === 1 ? 'page' : 'pages'} converted</strong>
        <p>{result.paragraphCount} paragraphs &middot; {result.tableCount} tables &middot; {result.imageCount} pictures</p>
        <details><summary>Conversion notes ({result.warnings.length})</summary>{result.warnings.map(warning => <p className="control-help" key={warning}>{warning}</p>)}</details>
      </section>}
    </div></div>
    <div className="assembly-actions"><button onClick={() => {
      abort.current?.abort();
      onClose();
    }}>{busy ? 'Cancel export' : 'Close'}</button>
      {!result && <button className="primary-action" disabled={busy} onClick={() => void convert()}>{busy ? <LoaderCircle size={16} /> : <FileText size={16} />}{busy ? 'Converting...' : 'Convert to Word'}</button>}
      {result && <button className="primary-action" onClick={() => {
        const bytes = Uint8Array.from(atob(result.base64), character => character.charCodeAt(0));
        const url = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
        const link = document.createElement('a');
        link.href = url;
        link.download = name.replace(/\.pdf$/i, '') + '.docx';
        link.click();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
      }}><Download size={16} />Download DOCX</button>}
    </div>
  </Modal>;
}
