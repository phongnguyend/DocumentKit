import { SheetColorDialog } from './components/SheetColorDialog';
import { ConditionalFormatDialog } from './components/ConditionalFormatDialog';
import { StrictMode, useEffect, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { Merge, Split, ImagePlus, Paintbrush, Calculator, Download, FolderOpen, Sheet as SheetIcon, Plus, X, Save, Trash2, Pencil, ArrowUp, ArrowDown, Undo2, Redo2, Bold, Italic } from 'lucide-react';
import { Modal } from './components/Modal';
import { ToolbarMenu } from './components/ToolbarMenu';
import { ImageSourceDialog } from './components/ImageSourceDialog';
import { columnPosition, rowPosition } from './types';
import { CellGrid } from './components/CellGrid';
import { address, cellValueType, cellEditValue, defaultStyle, inspect, parse, post, type Snapshot, type Style } from './types';
import './base.css';
import './style.css';
import './theme.css';
import { ThemeSwitcher } from './components/ThemeSwitcher';
import { applyTheme, readTheme } from './theme';

applyTheme(readTheme());

type Tab = { id: number; name: string; initial: Snapshot | null; dirty: boolean };

function Workbook({ tab, onDirty, onBusy }: { tab: Tab & { initial: Snapshot }; onDirty: (dirty: boolean) => void; onBusy: (busy: boolean) => void }) {
  const [snapshot, setSnapshot] = useState(tab.initial);
  const [past, setPast] = useState<Snapshot[]>([]);
  const [future, setFuture] = useState<Snapshot[]>([]);
  const [sheetId, setSheetId] = useState(tab.initial.info.sheets.find(sheet => !sheet.hidden)?.id ?? tab.initial.info.sheets[0].id);
  const [selected, setSelected] = useState('A1');
  const [selectedRange, setSelectedRange] = useState('A1');
  const [selectionAxis, setSelectionAxis] = useState<'row' | 'column' | null>(null);
  const [go, setGo] = useState('A1');
  const [value, setValue] = useState('');
  const [type, setType] = useState('auto');
  const [format, setFormat] = useState<Style>(defaultStyle);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [sheetColorOpen, setSheetColorOpen] = useState(false);
  const [conditionalOpen, setConditionalOpen] = useState(false);
  const [dialog, setDialog] = useState<string | null>(null);
  const [painter, setPainter] = useState<{ sheetId: string; address: string; locked: boolean } | null>(null);
  const [selectedImage, setSelectedImage] = useState<string>();
  const [imageDialog, setImageDialog] = useState<'add' | 'replace' | null>(null);
  const [sheetName, setSheetName] = useState('');
  const [rowHeight, setRowHeight] = useState(21);
  const [columnWidth, setColumnWidth] = useState(14);
  const sheet = snapshot.info.sheets.find(item => item.id === sheetId) ?? snapshot.info.sheets.find(item => !item.hidden)!;
  const image = sheet.images?.find(item => item.id === selectedImage);
  useEffect(() => {
    setSelectedImage(undefined);
  }, [sheet.id]);
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

  async function insertCells(rows: boolean, after: boolean) {
    const [first, last = first] = selectedRange.split(':').map(parse);
    const axis = rows ? 0 : 1;
    const position = after ? Math.max(first[axis], last[axis]) + 1 : Math.min(first[axis], last[axis]);
    const count = Math.abs(last[axis] - first[axis]) + 1;
    if (await edit({ action: rows ? 'insert-rows' : 'insert-columns', position, count })) {
      setSelected(address(rows ? position : row, rows ? column : position));
    }
  }

  async function deleteCells(rows: boolean) {
    const [first, last = first] = selectedRange.split(':').map(parse);
    const axis = rows ? 0 : 1;
    const position = Math.min(first[axis], last[axis]);
    const count = Math.abs(last[axis] - first[axis]) + 1;
    if (await edit({ action: rows ? 'delete-rows' : 'delete-columns', position, count })) {
      setSelected(address(rows ? position : row, rows ? column : position));
    }
  }

  async function visibility(rows: boolean, hidden: boolean, all = false) {
    const [first, last = first] = selectedRange.split(':').map(parse);
    const axis = rows ? 0 : 1;
    const position = all ? 1 : Math.min(first[axis], last[axis]);
    const count = all ? (rows ? 1048576 : 16384) : Math.abs(last[axis] - first[axis]) + 1;
    await edit({ action: (hidden ? 'hide-' : 'unhide-') + (rows ? 'rows' : 'columns'), position, count });
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
      <button disabled={busy} onClick={() => {
        setImageDialog('add');
        setPainter(null);
      }}><ImagePlus size={16} />Add image</button>
      <button disabled={busy || !!image || !selectedRange.includes(':') || selectedRange.split(':')[0] === selectedRange.split(':')[1]} onClick={async () => {
        if (await edit({ action: 'merge-cells', range: selectedRange })) {
          setSelected(selectedRange.split(':')[0]);
        }
      }}><Merge size={16} />Merge cells</button>
      <button disabled={busy || !!image || !sheet.merges.some(merge => {
        const [a, b = a] = merge.split(':').map(parse);
        const [c, d = c] = selectedRange.split(':').map(parse);
        return a[0] <= d[0] && b[0] >= c[0] && a[1] <= d[1] && b[1] >= c[1];
      })} onClick={() => void edit({ action: 'unmerge-cells', range: selectedRange })}><Split size={16} />Unmerge cells</button>
      <button aria-label="Format Painter" aria-pressed={!!painter} title="Format Painter: click for one cell, double-click to keep painting" disabled={busy} onClick={event => {
        if (event.detail < 2) {
          setPainter(current => current ? null : { sheetId: sheet.id, address: selected, locked: false });
        }
      }} onDoubleClick={() => setPainter({ sheetId: sheet.id, address: selected, locked: true })}><Paintbrush size={16} />Format Painter</button>
      <button disabled={busy} onClick={() => void edit({ action: 'recalculate' })}><Calculator size={16} />Recalculate</button>
      <button disabled={busy} onClick={() => setConditionalOpen(true)}><Paintbrush size={16} />Conditional formatting</button>
      <ToolbarMenu label="Insert" icon={Plus}>
        <button disabled={busy || !!image || Math.abs(parse(selectedRange.split(':').at(-1)!)[0] - parse(selectedRange.split(':')[0])[0]) >= 10000} onClick={() => void insertCells(true, false)}><Plus size={16} />Insert rows above</button>
        <button disabled={busy || !!image || Math.abs(parse(selectedRange.split(':').at(-1)!)[0] - parse(selectedRange.split(':')[0])[0]) >= 10000} onClick={() => void insertCells(true, true)}><Plus size={16} />Insert rows below</button>
        <button disabled={busy || !!image || Math.abs(parse(selectedRange.split(':').at(-1)!)[1] - parse(selectedRange.split(':')[0])[1]) >= 10000} onClick={() => void insertCells(false, false)}><Plus size={16} />Insert columns left</button>
        <button disabled={busy || !!image || Math.abs(parse(selectedRange.split(':').at(-1)!)[1] - parse(selectedRange.split(':')[0])[1]) >= 10000} onClick={() => void insertCells(false, true)}><Plus size={16} />Insert columns right</button>
      </ToolbarMenu>
      <ToolbarMenu label="Delete" icon={Trash2}>
        <button className="danger" disabled={busy || !!image || Math.abs(parse(selectedRange.split(':').at(-1)!)[0] - parse(selectedRange.split(':')[0])[0]) >= 10000} onClick={() => void deleteCells(true)}><Trash2 size={16} />Delete selected rows</button>
        <button className="danger" disabled={busy || !!image || Math.abs(parse(selectedRange.split(':').at(-1)!)[1] - parse(selectedRange.split(':')[0])[1]) >= 10000} onClick={() => void deleteCells(false)}><Trash2 size={16} />Delete selected columns</button>
      </ToolbarMenu>
      <ToolbarMenu label="Visibility" icon={SheetIcon} disabled={busy || !((selectionAxis && !image) || sheet.hiddenRows?.length || sheet.hiddenColumns?.length)}>
        {selectionAxis === 'row' && !image && <button disabled={busy} onClick={() => void visibility(true, true)}>Hide selected rows</button>}
        {selectionAxis === 'column' && !image && <button disabled={busy} onClick={() => void visibility(false, true)}>Hide selected columns</button>}
        {selectionAxis === 'row' && !image && <button disabled={busy} onClick={() => void visibility(true, false)}>Unhide selected rows</button>}
        {selectionAxis === 'column' && !image && <button disabled={busy} onClick={() => void visibility(false, false)}>Unhide selected columns</button>}
        {!!sheet.hiddenRows?.length && <button disabled={busy} onClick={() => void visibility(true, false, true)}>Unhide all rows</button>}
        {!!sheet.hiddenColumns?.length && <button disabled={busy} onClick={() => void visibility(false, false, true)}>Unhide all columns</button>}
      </ToolbarMenu>
      <ToolbarMenu label="Worksheet" icon={SheetIcon}>
        <button disabled={busy} onClick={() => setSheetColorOpen(true)}><Paintbrush size={16} />Worksheet color</button>
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
      <aside className="sheets-panel"><h2>Worksheets <small>{snapshot.info.sheets.length}</small></h2>{snapshot.info.sheets.map(item => <button key={item.id} style={{ boxShadow: item.tabColor ? `inset 4px 0 0 ${item.tabColor}` : undefined }} disabled={busy || item.hidden} aria-current={sheet.id === item.id ? 'page' : undefined} onClick={() => {
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
        <CellGrid onResize={(axis, address, size) => void edit({ action: 'resize-' + axis, cells: [{ address, value: '' }], ...(axis === 'row' ? { rowHeight: size } : { columnWidth: size }) })} onRangeChange={setSelectedRange} onSelectionAxisChange={setSelectionAxis} onClearRange={addresses => {
          if (addresses.length > 10000) {
            setError('Clear up to 10,000 cells at a time.');
            return;
          }
          void edit({ action: 'cells', cells: addresses.map(address => ({ address, value: '', type: 'clear' })) });
        }} date1904={snapshot.info.date1904} sheet={sheet} styles={snapshot.info.styles} selected={selected} onSelect={value => {
          setSelected(value);
          setSelectedImage(undefined);
        }} selectedImage={selectedImage} onSelectImage={id => {
          setSelectedImage(id);
          setPainter(null);
        }} onImageEdit={(imageId, action, bounds) => void edit({ action, imageId, ...bounds })} busy={busy} canEditValue={canEditValue && !painter} onPaint={painter ? destination => void paint(destination) : undefined} paintSource={painter?.sheetId === sheet.id ? painter.address : undefined} onEdit={text => void edit({ action: 'cells', cells: [cellEdit(text, selected, type)] })} onPaste={paste} />
        <div className="sheet-status"><select className="mobile-sheet-select" aria-label="Current worksheet" value={sheet.id} disabled={busy} onChange={event => {
          setSheetId(event.target.value);
          setSelected('A1');
        }}>{snapshot.info.sheets.filter(item => !item.hidden).map(item => <option key={item.id} value={item.id}>{item.name}</option>)}</select><span>{sheet.name} · {selected}</span><span>{busy ? 'Saving changes…' : 'Double-click to edit · Enter to save · Escape to cancel · Paste tab-separated cells'}</span></div>
      </main>
      <aside className="inspector">{image ? <><h2>Image <button aria-label="Close image inspector" onClick={() => setSelectedImage(undefined)}><X size={14} /></button></h2><p>{image.name}</p><p className="help">{Math.round(image.width)} x {Math.round(image.height)} px</p>
        <p className="help">{image.editable ? 'Drag the picture to move it. Drag a handle to resize. Release to save; Escape cancels.' : 'This picture has crop or rotation effects. Its preview is approximate; geometry editing is unavailable.'}</p>
        <div className="image-inspector-actions"><button disabled={busy} onClick={() => setImageDialog('replace')}><ImagePlus size={16} />Replace image</button><button onClick={() => {
          const link = document.createElement('a');
          link.href = image.dataUrl;
          link.download = image.name + (image.dataUrl.startsWith('data:image/png') ? '.png' : '.jpg');
          link.click();
        }}><Download size={16} />Download image</button><button className="danger" disabled={busy} onClick={() => void edit({ action: 'delete-image', imageId: image.id })}><Trash2 size={16} />Delete image</button></div></> : <><h2>Inspector <small>CELL</small></h2><div className="cell-summary"><SheetIcon size={20} /><strong>{selected}</strong><span>{sheet.name}</span></div>
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
      </>} </aside>
    </div>
    <footer><span>XlsxObjectKit · Native worksheet grid</span><span>Changes stay in this tab. Download to keep your workbook.</span></footer>
    {imageDialog && <ImageSourceDialog initialBounds={imageDialog === 'replace' && image ? { x: image.x, y: image.y, width: image.width, height: image.height } : { x: columnPosition(sheet, column), y: rowPosition(sheet, row), width: 240, height: 160 }} replacing={imageDialog === 'replace'} onClose={() => setImageDialog(null)} onApply={request => edit({ ...request, action: imageDialog === 'replace' ? 'replace-image' : 'add-image', imageId: image?.id })} />}
    {sheetColorOpen && <SheetColorDialog name={sheet.name} initialColor={sheet.tabColor} busy={busy} error={error} onClose={() => setSheetColorOpen(false)} onApply={tabColor => edit({ action: 'set-sheet-color', tabColor })} />}
    {conditionalOpen && <ConditionalFormatDialog sheet={sheet} initialRange={selectedRange} busy={busy} onEdit={edit} onClose={() => setConditionalOpen(false)} />}
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
    <header className="application-header"><div className="brand"><SheetIcon size={30} /><div><strong>XlsxObjectKit</strong><small>OBJECT PLAYGROUND</small></div></div><div className="header-actions"><ThemeSwitcher /><button disabled={busy} onClick={() => void open()}><Plus size={16} />New workbook</button><label className="upload-button"><FolderOpen size={16} />Open XLSX<input aria-label="Open XLSX" type="file" accept=".xlsx" multiple disabled={busy} onChange={event => {
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
