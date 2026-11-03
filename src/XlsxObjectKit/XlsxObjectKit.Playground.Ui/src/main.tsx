import { StrictMode, useEffect, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { Paintbrush, Calculator, Download, FolderOpen, Sheet as SheetIcon, Plus, X, Save, Trash2, Pencil, ArrowUp, ArrowDown, Undo2, Redo2, Bold, Italic } from 'lucide-react';
import { Modal } from './components/Modal';
import { ToolbarMenu } from './components/ToolbarMenu';
import { CellGrid } from './components/CellGrid';
import { address, cellValueType, cellEditValue, defaultStyle, inspect, parse, post, type Snapshot, type Style } from './types';
import './base.css';
import './style.css';

type Tab = { id: number; name: string; initial: Snapshot | null; dirty: boolean };

function Workbook({ tab, onDirty, onBusy }: { tab: Tab & { initial: Snapshot }; onDirty: (dirty: boolean) => void; onBusy: (busy: boolean) => void }) {
  const [snapshot, setSnapshot] = useState(tab.initial);
  const [past, setPast] = useState<Snapshot[]>([]);
  const [future, setFuture] = useState<Snapshot[]>([]);
  const [sheetId, setSheetId] = useState(tab.initial.info.sheets.find(sheet => !sheet.hidden)?.id ?? tab.initial.info.sheets[0].id);
  const [selected, setSelected] = useState('A1');
  const [go, setGo] = useState('A1');
  const [value, setValue] = useState('');
  const [type, setType] = useState('auto');
  const [format, setFormat] = useState<Style>(defaultStyle);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [dialog, setDialog] = useState<string | null>(null);
  const [painter, setPainter] = useState<{ sheetId: string; address: string; locked: boolean } | null>(null);
  const [sheetName, setSheetName] = useState('');
  const [rowHeight, setRowHeight] = useState(21);
  const [columnWidth, setColumnWidth] = useState(14);
  const sheet = snapshot.info.sheets.find(item => item.id === sheetId) ?? snapshot.info.sheets.find(item => !item.hidden)!;
  const cell = sheet.cells.find(item => item.address === selected);
  const [row, column] = parse(selected);
  const canEditValue = type !== 'date' && type !== 'error';
  useEffect(() => {
    setValue(cellEditValue(cell));
    setType(cellValueType(cell));
    setFormat(snapshot.info.styles[cell?.style ?? 0] ?? defaultStyle);
    setGo(selected);
    setRowHeight(sheet.rowHeights[row] ?? sheet.defaultRowHeight ?? 15);
    setColumnWidth(sheet.columnWidths[column] ?? sheet.defaultColumnWidth ?? 8.43);
  }, [snapshot, selected, sheet.id]);

  useEffect(() => {
    function cancelPainter(event: KeyboardEvent) {
      if (event.key === 'Escape') {
        setPainter(null);
      }
    }
    window.addEventListener('keydown', cancelPainter);
    return () => window.removeEventListener('keydown', cancelPainter);
  }, []);

  async function paint(destinations: string[]) {
    if (!painter || busy) {
      return;
    }
    if (destinations.length > 10000) {
      setError('Paint up to 10,000 cells at a time.');
      return;
    }
    const source = painter;
    const saved = await edit({ action: 'apply-format', sourceSheetId: source.sheetId, sourceAddress: source.address, addresses: destinations });
    if (saved) {
      setSelected(destinations[0]);
      if (!source.locked) {
        setPainter(current => current === source ? null : current);
      }
    }
  }

  async function edit(request: object) {
    if (busy) {
      return false;
    }
    if (!('action' in request) || request.action !== 'apply-format') {
      setPainter(null);
    }
    setBusy(true);
    onBusy(true);
    setError('');
    try {
      const response = await post('edit', snapshot.bytes, { sheetId: sheet.id, ...request });
      const next = await inspect(await response.blob());
      const diagnostics = response.headers.get('X-Calculation-Warnings');
      if (diagnostics) {
        next.info.warnings = [...JSON.parse(decodeURIComponent(diagnostics)), ...next.info.warnings];
      }
      setPast(previous => {
        const items = [...previous, snapshot].slice(-20);
        while (items.length > 1 && items.reduce((sum, item) => sum + item.bytes.size, 0) > 100 * 1024 * 1024) {
          items.shift();
        }
        return items;
      });
      setFuture([]);
      setSnapshot(next);
      onDirty(true);
      return true;
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : 'Edit failed.');
      return false;
    } finally {
      setBusy(false);
      onBusy(false);
    }
  }
  function cellEdit(text: string, at = selected, inputType = 'auto') {
    const kind = text === '' ? 'clear' : inputType !== 'auto' ? inputType : text.startsWith('=') ? 'formula' : /^(true|false)$/i.test(text) ? 'boolean' : /^-?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(text) ? 'number' : 'text';
    return { address: at, value: text, type: kind };
  }
  function paste(text: string) {
    try {
      const rows = text.replace(/\r/g, '').replace(/\n$/, '').split('\n');
      const cells = rows.flatMap((line, ri) => line.split('\t').map((item, ci) => cellEdit(item, address(row + ri, column + ci))));
      if (cells.length > 10000) {
        throw new Error('Paste at most 10,000 cells at a time.');
      }
      void edit({ action: 'cells', cells });
    } catch (cause) {
      setError(String(cause));
    }
  }
  function openSheetDialog(action: string) {
    setPainter(null);
    setSheetName(action === 'rename-sheet' ? sheet.name : 'Sheet' + (snapshot.info.sheets.length + 1));
    setDialog(action);
  }
  return <div className="workbook-shell">
    <div className="toolbar"><span className="document-name">{tab.name}</span><div className="toolbar-controls">
      <button aria-label="Undo" title="Undo" disabled={busy || !past.length} onClick={() => {
        setPainter(null);
        setSnapshot(past[past.length - 1]);
        setPast(past.slice(0, -1));
        setFuture([snapshot, ...future]);
        onDirty(true);
      }}><Undo2 size={16} /></button>
      <button aria-label="Redo" title="Redo" disabled={busy || !future.length} onClick={() => {
        setPainter(null);
        setSnapshot(future[0]);
        setPast([...past, snapshot]);
        setFuture(future.slice(1));
        onDirty(true);
      }}><Redo2 size={16} /></button>
      <button aria-label="Format Painter" aria-pressed={!!painter} title="Format Painter: click for one cell, double-click to keep painting" disabled={busy} onClick={event => {
        if (event.detail < 2) {
          setPainter(current => current ? null : { sheetId: sheet.id, address: selected, locked: false });
        }
      }} onDoubleClick={() => setPainter({ sheetId: sheet.id, address: selected, locked: true })}><Paintbrush size={16} />Format Painter</button>
      <button disabled={busy} onClick={() => void edit({ action: 'recalculate' })}><Calculator size={16} />Recalculate</button>
      <ToolbarMenu label="Worksheet" icon={SheetIcon}>
        <button disabled={busy} onClick={() => openSheetDialog('add-sheet')}><Plus size={16} />Add worksheet</button>
        <button disabled={busy} onClick={() => openSheetDialog('rename-sheet')}><Pencil size={16} />Rename worksheet</button>
        <button disabled={busy || snapshot.info.sheets.filter(item => !item.hidden).length <= 1} onClick={() => openSheetDialog('delete-sheet')}><Trash2 size={16} />Delete worksheet</button>
        <button disabled={busy || snapshot.info.sheets.indexOf(sheet) === 0} onClick={() => void edit({ action: 'move-sheet', position: snapshot.info.sheets.indexOf(sheet) - 1 })}><ArrowUp size={16} />Move worksheet left</button>
        <button disabled={busy || snapshot.info.sheets.indexOf(sheet) === snapshot.info.sheets.length - 1} onClick={() => void edit({ action: 'move-sheet', position: snapshot.info.sheets.indexOf(sheet) + 1 })}><ArrowDown size={16} />Move worksheet right</button>
      </ToolbarMenu>
      <button className="primary" disabled={busy} onClick={() => {
        const url = URL.createObjectURL(snapshot.bytes);
        const link = document.createElement('a');
        link.href = url;
        link.download = tab.name;
        link.click();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        onDirty(false);
      }}><Download size={16} />Download XLSX</button>
    </div></div>
    {error && <div role="alert" className="error-banner">{error}<button aria-label="Dismiss error" onClick={() => setError('')}><X size={14} /></button></div>}
    <div className="spreadsheet-layout">
      <aside className="sheets-panel"><h2>Worksheets <small>{snapshot.info.sheets.length}</small></h2>{snapshot.info.sheets.map(item => <button key={item.id} disabled={busy || item.hidden} aria-current={sheet.id === item.id ? 'page' : undefined} onClick={() => {
        setSheetId(item.id);
        setSelected('A1');
      }}><SheetIcon size={16} /><span>{item.name}</span>{item.hidden && <small>Hidden</small>}</button>)}<button disabled={busy} className="add-sheet" onClick={() => openSheetDialog('add-sheet')}><Plus size={16} />Add worksheet</button></aside>
      <main className="sheet-workspace">
    {painter && <div className="painter-status" role="status"><Paintbrush size={16} /><span>Copying format from {snapshot.info.sheets.find(item => item.id === painter.sheetId)?.name}!{painter.address}. {painter.locked ? 'Keep painting: click or drag across destination cells.' : 'Click or drag across destination cells.'} Escape to cancel.</span><button onClick={() => setPainter(null)} aria-label="Cancel Format Painter"><X size={16} /></button></div>}
        <form className="formula-bar" onSubmit={event => {
          event.preventDefault();
          void edit({ action: 'cells', cells: [cellEdit(value, selected, type)] });
        }}>
          <input aria-label="Go to cell" value={go} onChange={event => setGo(event.target.value)} onKeyDown={event => {
            if (event.key === 'Enter') {
              event.preventDefault();
              try {
                parse(go);
                setSelected(go.toUpperCase());
              } catch (cause) {
                setError(String(cause));
              }
            }
          }} /><span className="fx">fx</span><input aria-label="Formula bar" value={value} disabled={busy || !canEditValue} onChange={event => setValue(event.target.value)} /><button disabled={busy || !canEditValue} type="submit" title="Apply cell value" aria-label="Apply cell value"><Save size={16} /></button>
        </form>
        <CellGrid date1904={snapshot.info.date1904} sheet={sheet} styles={snapshot.info.styles} selected={selected} onSelect={setSelected} busy={busy} canEditValue={canEditValue && !painter} onPaint={painter ? destination => void paint(destination) : undefined} paintSource={painter?.sheetId === sheet.id ? painter.address : undefined} onEdit={text => void edit({ action: 'cells', cells: [cellEdit(text, selected, type)] })} onPaste={paste} />
        <div className="sheet-status"><select className="mobile-sheet-select" aria-label="Current worksheet" value={sheet.id} disabled={busy} onChange={event => {
          setSheetId(event.target.value);
          setSelected('A1');
        }}>{snapshot.info.sheets.filter(item => !item.hidden).map(item => <option key={item.id} value={item.id}>{item.name}</option>)}</select><span>{sheet.name} · {selected}</span><span>{busy ? 'Saving changes…' : 'Double-click to edit · Enter to save · Escape to cancel · Paste tab-separated cells'}</span></div>
      </main>
      <aside className="inspector"><h2>Inspector <small>CELL</small></h2><div className="cell-summary"><SheetIcon size={20} /><strong>{selected}</strong><span>{sheet.name}</span></div>
        <label>Value type<select value={type} disabled={busy} onChange={event => setType(event.target.value)}><option value="auto">Automatic</option><option value="text">Text</option><option value="number">Number</option><option value="boolean">Boolean</option><option value="formula">Formula</option>{type === 'date' && <option value="date" disabled>Date (read-only)</option>}{type === 'error' && <option value="error" disabled>Error (read-only)</option>}</select></label>
        <label>Value or formula<textarea value={value} disabled={busy || !canEditValue} onChange={event => setValue(event.target.value)} rows={4} /></label>
        <button className="primary" disabled={busy || !canEditValue} onClick={() => void edit({ action: 'cells', cells: [cellEdit(value, selected, type)] })}><Save size={16} />Apply value</button>
        {!canEditValue && <p className="help">Choose a supported value type to replace this value, or press Delete in the grid to clear it.</p>}
        {cell?.formula != null && <p className="help">{cell.hasCachedValue ? 'Calculated result: ' + (cell.type === 'b' ? cell.value === '1' ? 'TRUE' : 'FALSE' : cell.value || '(empty text)') : 'No calculated result. Recalculate and check Workbook notes for unsupported formulas or circular references.'}</p>}
        <section className="format-editor"><h2>Formatting</h2><div className="format-toggles"><button aria-label="Bold" aria-pressed={format.bold} onClick={() => setFormat({ ...format, bold: !format.bold })}><Bold size={16} /></button><button aria-label="Italic" aria-pressed={format.italic} onClick={() => setFormat({ ...format, italic: !format.italic })}><Italic size={16} /></button></div>
          <label>Text color<input type="color" value={format.color} onChange={event => setFormat({ ...format, color: event.target.value })} /></label><label>Background<input type="color" value={format.background} onChange={event => setFormat({ ...format, background: event.target.value })} /></label>
          <label>Alignment<select value={format.alignment} onChange={event => setFormat({ ...format, alignment: event.target.value as Style['alignment'] })}><option>general</option><option>left</option><option>center</option><option>right</option><option>justify</option><option>distributed</option><option>fill</option><option>centerContinuous</option></select></label>
          <button disabled={busy} onClick={() => void edit({ action: 'cells', cells: [{ address: selected, value: '' }], format })}><Save size={16} />Apply formatting</button>
        </section>
        <section className="format-editor"><h2>Row and column size</h2><label>Row height (pt)<input type="number" min="1" max="409" value={rowHeight} onChange={event => setRowHeight(+event.target.value)} /></label><label>Column width (characters)<input type="number" min="1" max="255" value={columnWidth} onChange={event => setColumnWidth(+event.target.value)} /></label><button disabled={busy} onClick={() => void edit({ action: 'dimensions', cells: [{ address: selected, value: '' }], rowHeight, columnWidth })}><Save size={16} />Apply sizes</button></section>
        <details open={snapshot.info.warnings.some(warning => warning.includes('!'))}><summary>Workbook notes</summary>{snapshot.info.warnings.map(warning => <p className="help" key={warning}>{warning}</p>)}</details>
      </aside>
    </div>
    <footer><span>XlsxObjectKit · Native worksheet grid</span><span>Changes stay in this tab. Download to keep your workbook.</span></footer>
    {dialog && <Modal labelledBy="sheet-dialog" onClose={() => setDialog(null)} canClose={!busy}><div className="modal-heading"><h2 id="sheet-dialog">{dialog === 'delete-sheet' ? 'Delete worksheet' : dialog === 'rename-sheet' ? 'Rename worksheet' : 'Add worksheet'}</h2><button aria-label="Close worksheet dialog" disabled={busy} onClick={() => setDialog(null)}><X size={18} /></button></div>
      {dialog === 'delete-sheet' ? <p>Delete “{sheet.name}”? You can undo this change in the current tab.</p> : <label>Worksheet name<input aria-label="Worksheet name" value={sheetName} maxLength={31} onChange={event => setSheetName(event.target.value)} /></label>}
      <div className="modal-actions"><button disabled={busy} onClick={() => setDialog(null)}>Cancel</button><button className="primary" disabled={busy} onClick={async () => {
        if (await edit({ action: dialog, name: sheetName })) {
          setDialog(null);
        }
      }}>{dialog === 'delete-sheet' ? 'Delete worksheet' : 'Save worksheet'}</button></div>{error && <p role="alert">{error}</p>}
    </Modal>}
  </div>;
}

function App() {
  const [tabs, setTabs] = useState<Tab[]>([{ id: 1, name: 'New tab', initial: null, dirty: false }]);
  const [active, setActive] = useState(1);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [dragging, setDragging] = useState(false);
  const [close, setClose] = useState<Tab | null>(null);
  const ids = useRef(1);
  async function open(files?: FileList | File[]) {
    if (busy) {
      return;
    }
    setBusy(true);
    setError('');
    try {
      const inputs = files ? Array.from(files) : [new File([await (await post('new')).blob()], 'Workbook.xlsx')];
      let blankId = tabs.find(tab => tab.id === active && tab.initial === null)?.id;
      if (tabs.length + inputs.length - (blankId === undefined ? 0 : 1) > 12) {
        throw new Error('Open up to 12 workbooks at a time.');
      }
      for (const file of inputs) {
        if (!file.name.toLowerCase().endsWith('.xlsx') || file.size > 50 * 1024 * 1024) {
          throw new Error('Choose an .xlsx file up to 50 MB.');
        }
        const initial = await inspect(file);
        const id = blankId ?? ++ids.current;
        const opened = { id, name: file.name, initial, dirty: false };
        if (blankId !== undefined) {
          setTabs(items => items.map(tab => tab.id === id ? opened : tab));
          blankId = undefined;
        } else {
          setTabs(items => [...items, opened]);
        }
        setActive(id);
      }
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : 'Could not open workbook.');
    } finally {
      setBusy(false);
    }
  }
  function remove(tab: Tab) {
    const remaining = tabs.filter(item => item.id !== tab.id);
    if (!remaining.length) {
      remaining.push({ id: ++ids.current, name: 'New tab', initial: null, dirty: false });
      setError('');
    }
    setTabs(remaining);
    if (active === tab.id) {
      setActive(remaining.at(-1)?.id ?? 0);
    }
    setClose(null);
  }
  useEffect(() => {
    function warn(event: BeforeUnloadEvent) {
      if (tabs.some(tab => tab.dirty)) {
        event.preventDefault();
      }
    }
    window.addEventListener('beforeunload', warn);
    return () => window.removeEventListener('beforeunload', warn);
  }, [tabs]);
  return <div className={'tabbed-playground' + (dragging ? ' file-drag' : '')} onDragOver={event => {
    if (event.dataTransfer.types.includes('Files')) {
      event.preventDefault();
      setDragging(true);
    }
  }} onDragLeave={event => {
    if (!event.currentTarget.contains(event.relatedTarget as Node)) {
      setDragging(false);
    }
  }} onDrop={event => {
    event.preventDefault();
    setDragging(false);
    if (event.dataTransfer.files.length) {
      void open(event.dataTransfer.files);
    }
  }}>
    <header className="application-header"><div className="brand"><SheetIcon size={30} /><div><strong>XlsxObjectKit</strong><small>OBJECT PLAYGROUND</small></div></div><div className="header-actions"><button disabled={busy} onClick={() => void open()}><Plus size={16} />New workbook</button><label className="upload-button"><FolderOpen size={16} />Open XLSX<input aria-label="Open XLSX" type="file" accept=".xlsx" multiple disabled={busy} onChange={event => {
      if (event.target.files?.length) {
        void open(Array.from(event.target.files));
      }
      event.target.value = '';
    }} /></label></div></header>
    <div className="document-tabs" role="tablist" aria-label="Workbooks">{tabs.map(tab => <div className={'document-tab' + (tab.id === active ? ' active' : '')} key={tab.id}><button role="tab" aria-selected={tab.id === active} disabled={busy} onClick={() => setActive(tab.id)}><SheetIcon size={14} /><span>{tab.name}{tab.dirty ? ' *' : ''}</span></button><button className="close-tab" aria-label={'Close ' + tab.name} disabled={busy} onClick={() => tab.dirty ? setClose(tab) : remove(tab)}><X size={14} /></button></div>)}<button className="new-tab" aria-label="New tab" title="New tab" disabled={busy || tabs.length >= 12} onClick={() => {
      const id = ++ids.current;
      setTabs(items => [...items, { id, name: 'New tab', initial: null, dirty: false }]);
      setActive(id);
      setError('');
    }}><Plus size={18} /></button></div>
    {error && <div role="alert" className="error-banner">{error}<button aria-label="Dismiss upload error" onClick={() => setError('')}><X size={16} /></button></div>}
    {(!tabs.length || tabs.some(tab => tab.id === active && !tab.initial)) && <div className="welcome" role="tabpanel" aria-label="Upload workbook"><div className="welcome-card"><SheetIcon size={52} /><h1>Work with your spreadsheets</h1><p>Open an XLSX workbook or create one to start editing.</p><label className="upload-button"><FolderOpen size={16} />Choose XLSX file<input aria-label="Choose XLSX file" type="file" accept=".xlsx" multiple disabled={busy} onChange={event => {
      if (event.target.files?.length) {
        void open(Array.from(event.target.files));
      }
      event.target.value = '';
    }} /></label><p className="help">Drop .xlsx files here · Up to 50 MB each</p><button className="primary" disabled={busy} onClick={() => void open()}><Plus size={16} />Create workbook</button></div></div>}
    {tabs.filter(tab => tab.initial !== null).map(tab => <div className="workbook-panel" role="tabpanel" aria-label={tab.name} key={tab.id} hidden={tab.id !== active}><Workbook tab={{ ...tab, initial: tab.initial! }} onBusy={setBusy} onDirty={dirty => setTabs(items => items.map(item => item.id === tab.id ? { ...item, dirty } : item))} /></div>)}
    {busy && <div role="status" className="busy-status">Working…</div>}
    {close && <Modal labelledBy="close-title" onClose={() => setClose(null)}><h2 id="close-title">Close workbook?</h2><p>“{close.name}” has changes that have not been downloaded.</p><div className="modal-actions"><button onClick={() => setClose(null)}>Keep editing</button><button className="danger" onClick={() => remove(close)}>Close without downloading</button></div></Modal>}
  </div>;
}

createRoot(document.getElementById('root')!).render(<StrictMode><App /></StrictMode>);
