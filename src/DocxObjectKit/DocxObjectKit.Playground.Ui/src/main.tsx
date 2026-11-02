import { StrictMode, useEffect, useRef, useState, useCallback } from 'react';
import { createPortal } from 'react-dom';
import { createRoot } from 'react-dom/client';
import { Download, FileText, FolderOpen, BookOpen, RotateCcw, RotateCw, Save, Trash2, X, Copy, Plus, ChevronLeft, ChevronRight, MousePointer2, TextCursor, ImagePlus, ImageIcon } from 'lucide-react';
import { marked } from 'marked';
import DOMPurify from 'dompurify';
import { createDocxPreview } from './renderers/docxRenderer';
import { PdfExportDialog } from './components/PdfExportDialog';
import { Modal } from './components/Modal';
import { PageThumbnail } from './components/PageThumbnail';
import { ImageEditor } from './components/ImageEditor';
import { ImageTransform } from './components/ImageTransform';
import { ImageSourceDialog } from './components/ImageSourceDialog';
import { ToolbarMenu } from './components/ToolbarMenu';
import './style.css';

type Inspection = import('./renderers/docxRenderer').DocxInspection;
type Snapshot = { bytes: Blob; info: Inspection; deck: Awaited<ReturnType<typeof createDocxPreview>> };
async function post(path: string, bytes: Blob, edit?: object) {
  let body: Blob | FormData = bytes;
  if (edit) {
    const form = new FormData();
    form.append('file', bytes, 'document.docx');
    form.append('edit', JSON.stringify(edit));
    body = form;
  }
  const response = await fetch('/api/documents/' + path, { method: 'POST', headers: edit ? undefined : { 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }, body });
  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    throw new Error(error.detail ?? 'The request failed.');
  }
  return response;
}
async function inspect(bytes: Blob): Promise<Snapshot> {
  const info: Inspection = await (await post('inspect', bytes)).json();
  const deck = await createDocxPreview(info);
  return { bytes, info, deck };
}
function download(bytes: Blob, name: string) {
  const url = URL.createObjectURL(bytes);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = name;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function PageCanvas({ snapshot, page, zoom, mode, selected, onSelect, busy, onTransform }: { busy: boolean; onTransform: (id: string, bounds: { x: number; y: number; width: number; height: number }) => Promise<boolean>; snapshot: Snapshot; page: number; zoom: string; mode: 'text' | 'objects'; selected: string | null; onSelect: (id: string | null) => void }) {
  const host = useRef<HTMLDivElement>(null);
  const [stageHost, setStageHost] = useState<HTMLDivElement | null>(null);
  const image = snapshot.info.pages[page - 1]?.images?.find(item => item.id === selected);
  const textObject = snapshot.info.pages[page - 1]?.textObjects.find(item => item.id === selected);
  const transformObject = image ?? (textObject?.bounds ? { id: textObject.id, ...textObject.bounds } : undefined);
  const [error, setError] = useState('');
  useEffect(() => {
    const abort = new AbortController();
    setStageHost(null);
    setError('');
    const canvas = document.createElement('canvas');
    const factor = Math.min(2, 1800 / snapshot.deck.width, 1800 / snapshot.deck.height);
    canvas.width = Math.ceil(snapshot.deck.width * factor);
    canvas.height = Math.ceil(snapshot.deck.height * factor);
    canvas.setAttribute('role', 'img');
    canvas.setAttribute('aria-label', `Page ${page}`);
    const stage = document.createElement('div');
    stage.className = 'page-stage';
    stage.append(canvas);
    const layer = document.createElement('div');
    layer.className = 'text-layer';
    layer.style.width = canvas.width + 'px';
    layer.style.height = canvas.height + 'px';
    const observer = new ResizeObserver(() => {
      const container = host.current;
      if (!container || !container.clientWidth || !container.clientHeight) {
        return;
      }
      const scale = zoom === 'fit' ? Math.min(container.clientWidth / canvas.width, container.clientHeight / canvas.height) : Number(zoom) / 100 * 96 / 72 / factor;
      stage.style.width = `${canvas.width * scale}px`;
      stage.style.height = `${canvas.height * scale}px`;
      layer.style.transform = `scale(${scale})`;
    });
    void snapshot.deck.render(canvas, page, abort.signal, layer).then(() => {
      if (!abort.signal.aborted) {
        for (const span of layer.querySelectorAll<HTMLElement>('[data-object-id]')) {
          span.classList.toggle('selected-object', span.dataset.objectId === selected);
        }
        stage.append(layer);
        host.current?.replaceChildren(stage);
        setStageHost(stage);
        if (host.current) {
          observer.observe(host.current);
        }
      }
    }).catch(cause => {
      if (!abort.signal.aborted) {
        host.current?.replaceChildren();
        setError(cause.message);
      }
    });
    return () => {
      abort.abort();
      observer.disconnect();
    };
  }, [snapshot, page, selected, zoom]);
  return <>{error && <p role="alert">Preview unavailable: {error}</p>}<div ref={host} className={`canvas-host ${mode}-mode ${zoom === 'fit' ? '' : 'zoomed'}`} onClick={event => {
    if (mode === 'objects') {
      onSelect((event.target as HTMLElement).closest<HTMLElement>('[data-object-id]')?.dataset.objectId ?? null);
    }
  }} />{stageHost && mode === 'objects' && transformObject && !transformObject.grouped && transformObject.width > 0 && transformObject.height > 0 && createPortal(<ImageTransform kind={image ? 'image' : 'text'} key={transformObject.id} image={transformObject} pageWidth={snapshot.deck.width} pageHeight={snapshot.deck.height} busy={busy} onSave={bounds => onTransform(transformObject.id, bounds)} />, stageHost)}</>;
}

function MarkdownView({ snapshot, name, onClose }: { snapshot: Snapshot; name: string; onClose: () => void }) {
  const [images, setImages] = useState(true);
  const [source, setSource] = useState(false);
  const [markdown, setMarkdown] = useState('');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [copied, setCopied] = useState(false);
  const alive = useRef(true);
  async function convert() {
    setBusy(true);
    setError('');
    try {
      const data = await (await post(`markdown?images=${images}`, snapshot.bytes)).json();
      if (alive.current) {
        setMarkdown(data.markdown);
      }
    } catch (cause) {
      if (alive.current) {
        setError((cause as Error).message);
      }
    } finally {
      if (alive.current) {
        setBusy(false);
      }
    }
  }
  useEffect(() => {
    alive.current = true;
    void convert();
    return () => { alive.current = false; };
  }, []);
  return <Modal labelledBy="markdown-title" className="markdown-modal" onClose={onClose}>
    <div className="modal-heading"><h2 id="markdown-title">View Markdown</h2><button aria-label="Close Markdown" onClick={onClose}><X size={18} /></button></div>
    <div className="markdown-toolbar"><div className="mode-switch"><button aria-pressed={!source} onClick={() => setSource(false)}>Preview</button><button aria-pressed={source} onClick={() => setSource(true)}>Source</button></div><label><input type="checkbox" checked={images} onChange={event => setImages(event.target.checked)} />Images</label><button disabled={busy} onClick={() => void convert()}>{busy ? 'Converting...' : 'Reapply options'}</button></div>
    {error && <p role="alert">{error}</p>}
    <div className="markdown-body">{source ? <pre>{markdown}</pre> : <article dangerouslySetInnerHTML={{ __html: DOMPurify.sanitize(marked.parse(markdown, { async: false })) }} />}</div>
    <div className="modal-actions"><button disabled={!markdown} onClick={() => void navigator.clipboard.writeText(markdown).then(() => setCopied(true)).catch(() => setError('Clipboard unavailable. Use Source to select and copy the text.'))}><Copy size={16} />{copied ? 'Copied' : 'Copy Markdown'}</button><button disabled={!markdown} onClick={() => download(new Blob([markdown], { type: 'text/markdown' }), name.replace(/\.docx$/i, '') + '.md')}><Download size={16} />Download Markdown</button></div>
  </Modal>;
}

type OpenedDocument = { snapshot: Snapshot; name: string };
type PageProps = { active: boolean; headerHost: HTMLDivElement | null; id: string; initial?: OpenedDocument; onOpen: (id: string, documents: OpenedDocument[]) => void; onStatus: (id: string, dirty: boolean, locked: boolean) => void };
function App({ id, initial, onOpen, onStatus, active, headerHost }: PageProps) {
  const [current, setCurrent] = useState<Snapshot | null>(initial?.snapshot ?? null);
  const name = initial?.name ?? '';
  const [saved, setSaved] = useState<Blob | undefined>(initial?.snapshot.bytes);
  const [page, setPage] = useState(1);
  const [zoom, setZoom] = useState('fit');
  const [addingImage, setAddingImage] = useState(false);
  const [mode, setMode] = useState<'text' | 'objects'>('text');
  const [selected, setSelected] = useState<string | null>(null);
  const [draft, setDraft] = useState('');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [dragging, setDragging] = useState(false);
  const [history, setHistory] = useState<Snapshot[]>([]);
  const [future, setFuture] = useState<Snapshot[]>([]);
  const [markdown, setMarkdown] = useState(false);
  const [pdfExport, setPdfExport] = useState(false);
  const fileInput = useRef<HTMLInputElement>(null);
  const currentPage = current?.info.pages[page - 1];
  const selectedImage = currentPage?.images?.find(item => item.id === selected);
  const object = currentPage?.textObjects.find(item => item.id === selected);
  useEffect(() => {
    setSelected(null);
    setDraft('');
  }, [page, current]);

  useEffect(() => {
    onStatus(id, !!current && current.bytes !== saved, busy || markdown || pdfExport || addingImage);
  }, [id, current, saved, busy, markdown, pdfExport, addingImage, onStatus]);

  async function open(files: File[]) {
    if (!files.length || busy) {
      return;
    }
    if (files.some(file => !/\.docx$/i.test(file.name) || file.size === 0 || file.size > 50 * 1024 * 1024)) {
      setError('Choose non-empty .docx files up to 50 MB each.');
      return;
    }
    setBusy(true);
    setError('');
    try {
      const documents: OpenedDocument[] = [];
      for (const file of files) {
        documents.push({ snapshot: await inspect(file), name: file.name });
      }
      onOpen(id, documents);
    } catch (cause) {
      setError((cause as Error).message);
    } finally {
      setBusy(false);
    }
  }
  async function edit(request: object, nextPage = page) {
    if (!current || busy) {
      return false;
    }
    setBusy(true);
    setError('');
    try {
      const output = await (await post('edit', current.bytes, request)).blob();
      const updated = await inspect(output);
      setHistory(items => [...items.slice(-9), current]);
      setFuture([]);
      setCurrent(updated);
      setPage(Math.min(nextPage, updated.info.pages.length));
      return true;
    } catch (cause) {
      setError((cause as Error).message);
      return false;
    } finally {
      setBusy(false);
    }
  }
  return <div className="app-shell">
    {active && headerHost && createPortal(<><button disabled={busy} onClick={() => fileInput.current?.click()}><FolderOpen size={17} />{busy ? 'Opening...' : 'Open DOCX'}</button><input ref={fileInput} type="file" multiple hidden accept=".docx" aria-label="Open DOCX" onChange={event => { void open(Array.from(event.target.files ?? [])); event.target.value = ''; }} /></>, headerHost)}
    {error && <div role="alert" className="error-banner">{error}<button aria-label="Dismiss error" onClick={() => setError('')}><X size={16} /></button></div>}
    {busy && <p role="status" className="status">Processing document...</p>}
    {!current ? <main className="welcome"><div className={`dropzone ${dragging ? 'dragging' : ''}`} onDragOver={event => { event.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={event => { event.preventDefault(); setDragging(false); void open(Array.from(event.dataTransfer.files)); }}><BookOpen size={48} /><h1>Your documents, ready to work with</h1><p>Open a document to preview pages, edit text, manage images, and extract Markdown.</p><button className="primary" disabled={busy} onClick={() => fileInput.current?.click()}>Choose document</button><small>Drop a .docx here Ã‚· Up to 50 MB</small></div></main> : <>
      <div className="toolbar document-bar"><strong className="filename" title={name}>{name}</strong><span>{current.info.pages.length} pages</span><div className="toolbar-controls"><div className="interaction-mode" role="group" aria-label="Interaction mode">{([['objects', 'Select objects', MousePointer2], ['text', 'Select text', TextCursor]] as const).map(([value, label, Icon]) => <button key={value} title={label} aria-label={label} aria-pressed={mode === value} onClick={() => { setMode(value); window.getSelection()?.removeAllRanges(); }}><Icon size={18} /></button>)}</div>
        <button disabled={busy} onClick={() => void edit({ kind: 'add-text', pageNumber: page, text: 'New text' })}><Plus size={16} />Add text</button><button disabled={busy} onClick={() => setAddingImage(true)}><ImagePlus size={16} />Add image</button>
        <ToolbarMenu label="Document" icon={FileText}><button disabled={busy} onClick={() => setPdfExport(true)}><Download size={16} />Export PDF</button><button disabled={busy} onClick={() => setMarkdown(true)}><FileText size={16} />View Markdown</button></ToolbarMenu>
        <button className="toolbar-primary-download" disabled={busy} onClick={() => { download(current.bytes, name); setSaved(current.bytes); }}><Download size={16} />Download DOCX</button>
        <button aria-label="Undo" disabled={busy || !history.length} onClick={() => { setFuture(items => [...items, current]); setCurrent(history[history.length - 1]); setHistory(items => items.slice(0, -1)); setPage(1); }}><RotateCcw size={16} /></button><button aria-label="Redo" disabled={busy || !future.length} onClick={() => { setHistory(items => [...items, current]); setCurrent(future[future.length - 1]); setFuture(items => items.slice(0, -1)); setPage(1); }}><RotateCw size={16} /></button>
      </div></div>
      <main className="workspace"><aside className="pages-panel"><h2>Pages</h2><p className="help">Pages flow automatically as content changes.</p><ol>{current.info.pages.map(item => <li key={item.id}><button className="page-card" disabled={busy} aria-label={`Page ${item.number}: ${item.title || 'Untitled'}`} title={item.title} aria-current={page === item.number ? 'page' : undefined} onClick={() => setPage(item.number)}>
          <PageThumbnail deck={current.deck} page={item.number} />
          <span className="page-caption"><span className="page-number">{item.number}</span><span className="page-title">{item.title || 'Untitled page'}</span></span>
        </button></li>)}</ol></aside>
      <section className="page-workspace" aria-label="Page workspace"><div className="canvas-toolbar"><span>Page {page} of {current.info.pages.length}</span></div>{currentPage && <PageCanvas busy={busy} onTransform={(id, bounds) => edit({ kind: currentPage.textObjects.some(item => item.id === id) ? 'transform-text' : 'resize-image', pageNumber: page, objectId: id, ...bounds })} snapshot={current} page={page} zoom={zoom} mode={mode} selected={selected} onSelect={id => {
        const item = currentPage.textObjects.find(candidate => candidate.id === id) ?? currentPage.images?.find(candidate => candidate.id === id);
        setSelected(item?.id ?? null);
        setDraft(item && 'text' in item ? item.text : '');
      }} />}<details className="render-notes"><summary>Preview limitations</summary>{current.info.warnings.map(note => <p key={note}>{note}</p>)}</details></section>
      <aside className="inspector"><h2>Inspector</h2><p className="help">Use Select objects to click text on the page and edit it, or choose an object below. Images support replacement, download, resizing and deletion.</p><div className="object-list">{currentPage?.textObjects.map(item => <button key={item.id} aria-pressed={selected === item.id} disabled={busy} onClick={() => { setSelected(item.id); setDraft(item.text); }}><FileText size={16} /><span>{item.text || item.name}</span></button>)}{currentPage?.images?.filter(item => !item.readOnly).map(item => <button key={item.id} aria-pressed={selected === item.id} disabled={busy} onClick={() => { setMode('objects'); setSelected(item.id); }}><ImageIcon size={16} /><span>{item.name}</span></button>)}</div>{selectedImage && current && <ImageEditor page={page} loadImage={() => current.deck.image(selectedImage.partName)} key={selectedImage.id} image={selectedImage} busy={busy} onEdit={request => edit({ ...request, pageNumber: page, objectId: selectedImage.id })} onDownload={async () => download(await current.deck.image(selectedImage.partName), selectedImage.partName.split('/').pop() ?? 'image')} />}{object && <div className="text-editor"><p className="help">{object.bounds?.grouped ? 'Table text and paragraphs spanning multiple pages can be edited, but cannot be positioned independently.' : 'Drag to create a page-positioned text frame; resize to change wrapping. Release to save. Escape cancels.'}</p><label>Text<textarea aria-label="Text" maxLength={4000} value={draft} onChange={event => setDraft(event.target.value)} disabled={busy} rows={8} /></label><p className="help">Replacement keeps the first run style in each paragraph. Mixed formatting within a paragraph is simplified.</p><button className="primary" disabled={busy || draft === object.text} onClick={() => void edit({ kind: 'text', pageNumber: page, objectId: object.id, text: draft })}><Save size={16} />Save text</button><button disabled={busy} onClick={() => void edit({ kind: 'delete-text', objectId: object.id })}><Trash2 size={16} />Delete paragraph</button></div>}</aside></main>
    </>}
    {current && <div className="page-navigation-bar" role="toolbar" aria-label="Page navigation"><button aria-label="Previous page" disabled={busy || page === 1} onClick={() => setPage(page - 1)}><ChevronLeft size={18} /></button><label>Page <select aria-label="Page" disabled={busy} value={page} onChange={event => setPage(Number(event.target.value))}>{current.info.pages.map(item => <option key={item.id} value={item.number}>{item.number}</option>)}</select> / {current.info.pages.length}</label><button aria-label="Next page" disabled={busy || page === current.info.pages.length} onClick={() => setPage(page + 1)}><ChevronRight size={18} /></button><label>Zoom <select aria-label="Zoom" value={zoom} onChange={event => setZoom(event.target.value)}><option value="fit">Fit page</option>{[50, 75, 100, 125, 150, 200].map(value => <option key={value} value={value}>{value}%</option>)}</select></label></div>}
    <footer>Experimental browser preview<span>Edits stay in this session. Download to keep your changes.</span></footer>
    {pdfExport && current && <PdfExportDialog bytes={current.bytes} name={name} page={page} count={current.info.pages.length} onClose={() => setPdfExport(false)} />}
    {markdown && current && <MarkdownView snapshot={current} name={name} onClose={() => setMarkdown(false)} />}
    {addingImage && <ImageSourceDialog initialBounds={{ x: 36, y: 36, width: 180, height: 120 }} onClose={() => setAddingImage(false)} onApply={request => edit({ ...request, kind: 'add-image', pageNumber: page })} />}
  </div>;
}

type Tab = { id: string; document?: OpenedDocument; dirty?: boolean; locked?: boolean };
function DocumentTabs() {
  const [headerHost, setHeaderHost] = useState<HTMLDivElement | null>(null);
  const [tabs, setTabs] = useState<Tab[]>(() => [{ id: crypto.randomUUID() }]);
  const [active, setActive] = useState('');
  const [closing, setClosing] = useState<string | null>(null);
  const [dragged, setDragged] = useState<string | null>(null);
  const current = tabs.find(tab => tab.id === active) ?? tabs[0];
  const locked = !!current.locked || closing !== null;
  const status = useCallback((id: string, dirty: boolean, busy: boolean) => {
    setTabs(items => items.map(tab => tab.id === id && (tab.dirty !== dirty || tab.locked !== busy) ? { ...tab, dirty, locked: busy } : tab));
  }, []);
  const open = useCallback((id: string, documents: OpenedDocument[]) => {
    const added = documents.map(document => ({ id: crypto.randomUUID(), document }));
    setTabs(items => [...items.filter(tab => tab.id !== id || tab.document), ...added]);
    setActive(added[0].id);
  }, []);
  useEffect(() => {
    if (!tabs.some(tab => tab.dirty)) {
      return;
    }
    const warn = (event: BeforeUnloadEvent) => {
      event.preventDefault();
      event.returnValue = '';
    };
    window.addEventListener('beforeunload', warn);
    return () => window.removeEventListener('beforeunload', warn);
  }, [tabs]);
  function remove(id: string) {
    const index = tabs.findIndex(tab => tab.id === id);
    const next = tabs.filter(tab => tab.id !== id);
    if (!next.length) {
      next.push({ id: crypto.randomUUID() });
    }
    setTabs(next);
    if (current.id === id) {
      setActive(next[Math.min(index, next.length - 1)].id);
    }
    setClosing(null);
  }
  return <div className="tabbed-playground">
    <header className="application-header"><div className="brand"><span className="logo">D</span><div><strong>DocxObjectKit</strong><span className="subtitle">OBJECT PLAYGROUND</span></div></div><div className="app-header-actions" ref={setHeaderHost} /></header>
    <div className="document-tabs" role="tablist" aria-label="Documents">
      {tabs.map((tab, index) => <div className={`document-tab ${current.id === tab.id ? 'active' : ''}`} key={tab.id} draggable={!locked} onDragStart={() => setDragged(tab.id)} onDragEnd={() => setDragged(null)} onDragOver={event => event.preventDefault()} onDrop={event => {
        event.preventDefault();
        if (locked || !dragged || dragged === tab.id) {
          return;
        }
        const moved = tabs.find(item => item.id === dragged);
        if (moved) {
          const order = tabs.filter(item => item.id !== dragged);
          order.splice(index, 0, moved);
          setTabs(order);
        }
        setDragged(null);
      }}>
        <button role="tab" id={`tab-${tab.id}`} aria-controls={`panel-${tab.id}`} aria-selected={current.id === tab.id} tabIndex={current.id === tab.id ? 0 : -1} disabled={locked} title={tab.document?.name ?? 'New document'} onClick={() => setActive(tab.id)} onKeyDown={event => {
          let destination = index;
          if (event.key === 'ArrowRight') {
            destination = (index + 1) % tabs.length;
          } else if (event.key === 'ArrowLeft') {
            destination = (index + tabs.length - 1) % tabs.length;
          } else if (event.key === 'Home') {
            destination = 0;
          } else if (event.key === 'End') {
            destination = tabs.length - 1;
          } else {
            return;
          }
          event.preventDefault();
          setActive(tabs[destination].id);
          document.getElementById(`tab-${tabs[destination].id}`)?.focus();
        }}><FileText size={15} /><span>{tab.document?.name ?? 'New document'}</span>{tab.dirty && <span className="modified-dot" aria-label="Unsaved changes">{String.fromCharCode(8226)}</span>}</button>
        <button className="close-tab" disabled={locked || tab.locked} aria-label={`Close ${tab.document?.name ?? 'new document'}`} onClick={() => tab.dirty ? setClosing(tab.id) : remove(tab.id)}><X size={14} /></button>
      </div>)}
      <button className="new-tab" disabled={locked} aria-label="New tab" onClick={() => {
        const tab = { id: crypto.randomUUID() };
        setTabs(items => [...items, tab]);
        setActive(tab.id);
      }}><Plus size={18} /></button>
    </div>
    {tabs.map(tab => <div key={tab.id} role="tabpanel" id={`panel-${tab.id}`} aria-labelledby={`tab-${tab.id}`} hidden={current.id !== tab.id} className="document-panel"><App active={current.id === tab.id} headerHost={headerHost} id={tab.id} initial={tab.document} onOpen={open} onStatus={status} /></div>)}
    {closing && <Modal labelledBy="close-tab-title" className="confirm-modal" onClose={() => setClosing(null)}><div className="modal-heading"><h2 id="close-tab-title">Close document?</h2><button aria-label="Cancel close" onClick={() => setClosing(null)}><X size={18} /></button></div><p>Changes in {tabs.find(tab => tab.id === closing)?.document?.name} have not been downloaded. Closing this tab discards them.</p><div className="modal-actions"><button onClick={() => setClosing(null)}>Cancel</button><button className="danger" onClick={() => remove(closing)}>Discard and close</button></div></Modal>}
  </div>;
}
createRoot(document.getElementById('root')!).render(<StrictMode><DocumentTabs /></StrictMode>);
