import { useEffect, useMemo, useRef, useState } from 'react';
import { WorksheetImage } from './WorksheetImage';
import { columnPosition, rowPosition } from '../types';
import { cellAppearance, displayValue, textAppearance } from '../formatting';
import { address, cellEditValue, defaultStyle, parse, type Sheet, type Style } from '../types';

export function CellGrid({ sheet, styles, selected, onSelect, onEdit, onPaste, busy, date1904, canEditValue = true, onPaint, paintSource, selectedImage, onSelectImage, onImageEdit }: {
  selectedImage?: string; onSelectImage?: (id: string) => void; onImageEdit?: (id: string, action: string, bounds?: object) => void;
  sheet: Sheet; styles: Style[]; selected: string; onSelect: (value: string) => void;
  onEdit: (value: string) => void; onPaste: (value: string) => void; busy: boolean; date1904?: boolean; canEditValue?: boolean; onPaint?: (addresses: string[]) => void; paintSource?: string;
}) {
  const [scroll, setScroll] = useState(0);
  const [viewportHeight, setViewportHeight] = useState(700);
  const [editing, setEditing] = useState<string | null>(null);
  const [paintRange, setPaintRange] = useState<{ start: [number, number]; end: [number, number] } | null>(null);
  const drag = useRef<typeof paintRange>(null);
  const suppressClick = useRef(false);
  const host = useRef<HTMLDivElement>(null);
  const [selectedRow, selectedColumn] = parse(selected);
  // Bound the browser scroll area; Go to cell opens any other worksheet window.
  const base = Math.floor((selectedRow - 1) / 10000) * 10000;
  const firstColumn = Math.floor((selectedColumn - 1) / 26) * 26 + 1;
  const columns = Array.from({ length: Math.min(26, 16385 - firstColumn) }, (_, i) => firstColumn + i);
  const count = Math.min(10000, 1048576 - base);
  const cells = useMemo(() => new Map(sheet.cells.map(cell => [cell.address, cell])), [sheet]);
  const widths = columns.map(column => Math.max(1, Math.min(255, sheet.columnWidths[column] ?? sheet.defaultColumnWidth ?? 8.43)) * 7 + 5);
  const x = [48];
  widths.forEach(width => x.push(x[x.length - 1] + width));
  const y = useMemo(() => {
    const offsets = [32];
    for (let i = 1; i <= count; i++) {
      offsets.push(offsets[i - 1] + Math.max(1, Math.min(409, sheet.rowHeights[base + i] ?? sheet.defaultRowHeight ?? 15)) * 4 / 3);
    }
    return offsets;
  }, [sheet, base, count]);
  const start = Math.max(0, y.findIndex((offset, index) => index < count && y[index + 1] > scroll) - 3);
  const lastVisible = y.findIndex(offset => offset >= scroll + viewportHeight);
  const end = Math.min(count, Math.max(start + 1, (lastVisible < 0 ? count : lastVisible) + 4));
  const merges = useMemo(() => sheet.merges.map(range => {
    const ends = range.split(':');
    return { first: parse(ends[0]), last: parse(ends.at(-1)!) };
  }), [sheet]);
  useEffect(() => {
    const observer = new ResizeObserver(entries => {
      const height = entries[0]?.contentRect.height;
      if (height > 0) {
        setViewportHeight(height);
      }
    });
    if (host.current) {
      observer.observe(host.current);
    }
    return () => observer.disconnect();
  }, []);
  useEffect(() => {
    setEditing(null);
    const top = y[selectedRow - base - 1];
    if (host.current && (top < host.current.scrollTop + 32 || top > host.current.scrollTop + host.current.clientHeight - 40)) {
      host.current.scrollTop = Math.max(0, top - 50);
    }
  }, [selected, sheet.id]);

  function cancelDrag() {
    drag.current = null;
    setPaintRange(null);
  }
  useEffect(() => {
    cancelDrag();
  }, [!!onPaint, sheet.id]);
  useEffect(() => {
    function escape(event: KeyboardEvent) {
      if (event.key === 'Escape') {
        cancelDrag();
      }
    }
    window.addEventListener('keydown', escape);
    return () => window.removeEventListener('keydown', escape);
  }, []);
  function updateDrag(clientX: number, clientY: number) {
    if (!drag.current || !host.current) {
      return;
    }
    const rect = host.current.getBoundingClientRect();
    const px = clientX - rect.left + host.current.scrollLeft;
    const py = clientY - rect.top + host.current.scrollTop;
    const ci = Math.max(0, Math.min(columns.length - 1, x.findIndex((value, i) => i < columns.length && x[i + 1] > px)));
    const ri = py >= y[count] ? count - 1 : Math.max(0, y.findIndex((value, i) => i < count && y[i + 1] > py));
    const column = px >= x[columns.length] ? columns.at(-1)! : columns[ci];
    drag.current = { ...drag.current, end: [base + ri + 1, column] };
    setPaintRange(drag.current);
  }
  const rangeTop = paintRange ? Math.min(paintRange.start[0], paintRange.end[0]) : 0;
  const rangeBottom = paintRange ? Math.max(paintRange.start[0], paintRange.end[0]) : 0;
  const rangeLeft = paintRange ? Math.min(paintRange.start[1], paintRange.end[1]) : 0;
  const rangeRight = paintRange ? Math.max(paintRange.start[1], paintRange.end[1]) : 0;

  return <div className={'cell-grid' + (onPaint ? ' painting' : '')} ref={host} role="grid" aria-label={sheet.name + ' cells'} aria-rowcount={1048576} aria-colcount={16384} tabIndex={0}
    onPointerMove={event => updateDrag(event.clientX, event.clientY)}
    onPointerCancel={cancelDrag} onLostPointerCapture={cancelDrag}
    onPointerUp={event => {
      if (!drag.current || !onPaint || busy) {
        return;
      }
      updateDrag(event.clientX, event.clientY);
      const range = drag.current;
      const targets: string[] = [];
      for (let row = Math.min(range.start[0], range.end[0]); row <= Math.max(range.start[0], range.end[0]) && targets.length <= 10000; row++) {
        for (let column = Math.min(range.start[1], range.end[1]); column <= Math.max(range.start[1], range.end[1]) && targets.length <= 10000; column++) {
          const merge = merges.find(item => row >= item.first[0] && row <= item.last[0] && column >= item.first[1] && column <= item.last[1]);
          if (!merge || (row === merge.first[0] && column === merge.first[1])) {
            targets.push(address(row, column));
          }
        }
      }
      cancelDrag();
      if (event.currentTarget.hasPointerCapture(event.pointerId)) {
        event.currentTarget.releasePointerCapture(event.pointerId);
      }
      if (targets.length > 0) {
        onPaint(targets);
      }
    }}
    onScroll={event => setScroll(event.currentTarget.scrollTop)} onCopy={event => {
      if (editing !== null) {
        return;
      }
      event.clipboardData.setData('text/plain', cells.get(selected)?.value ?? '');
      event.preventDefault();
    }} onPaste={event => {
      if (busy || editing !== null || onPaint) {
        return;
      }
      event.preventDefault();
      onPaste(event.clipboardData.getData('text/plain'));
    }} onKeyDown={event => {
      if (busy || editing !== null) {
        return;
      }
      const delta = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowRight: [0, 1], ArrowLeft: [0, -1] }[event.key];
      if (delta) {
        event.preventDefault();
        onSelect(address(Math.max(1, Math.min(1048576, selectedRow + delta[0])), Math.max(1, Math.min(16384, selectedColumn + delta[1]))));
      } else if (onPaint && event.key === 'Enter') {
        event.preventDefault();
        onPaint([selected]);
      } else if (canEditValue && (event.key === 'Enter' || event.key === 'F2')) {
        event.preventDefault();
        const cell = cells.get(selected);
        setEditing(cellEditValue(cell));
      } else if (!onPaint && event.key === 'Delete') {
        event.preventDefault();
        onEdit('');
      }
    }}>
    <div className="grid-surface" style={{ width: x.at(-1), height: y.at(-1) }}>
      {paintRange && <div className="paint-range" aria-hidden="true" style={{ left: x[rangeLeft - firstColumn], top: y[rangeTop - base - 1], width: x[rangeRight - firstColumn + 1] - x[rangeLeft - firstColumn], height: y[rangeBottom - base] - y[rangeTop - base - 1] }} />}
      <div className="column-headings" role="row" style={{ width: x.at(-1) }}><div className="corner" />{columns.map((column, i) => <div role="columnheader" key={column} style={{ left: x[i], width: widths[i] }}>{address(1, column).replace('1', '')}</div>)}</div>
      {Array.from({ length: end - start }, (_, i) => i + start).map(index => <div key={index} role="row" aria-rowindex={base + index + 1}>
        <div className="row-heading" role="rowheader" style={{ top: y[index], height: y[index + 1] - y[index] }}>{base + index + 1}</div>
        {columns.map((column, ci) => {
          const row = base + index + 1;
          const key = address(row, column);
          const merge = merges.find(item => row >= item.first[0] && row <= item.last[0] && column >= item.first[1] && column <= item.last[1]);
          if (merge && (row !== merge.first[0] || column !== merge.first[1])) {
            return null;
          }
          const cell = cells.get(key);
          const style = styles[cell?.style ?? sheet.rowStyles?.[row] ?? sheet.columnStyles?.[column] ?? 0] ?? defaultStyle;
          const width = merge ? x[Math.min(columns.length, merge.last[1] - firstColumn + 1)] - x[ci] : widths[ci];
          const height = merge ? y[Math.min(count, merge.last[0] - base)] - y[index] : y[index + 1] - y[index];
          return <div key={column} role="gridcell" aria-label={key} aria-selected={selected === key} aria-colindex={column}
            className={'grid-cell' + (selected === key ? ' selected' : '') + (paintSource === key ? ' paint-source' : '')} style={{ top: y[index], left: x[ci], width, height, ...cellAppearance(style, cell) }}
            onPointerDown={event => {
              suppressClick.current = false;
              if (onPaint && !busy && event.button === 0) {
                event.preventDefault();
                suppressClick.current = true;
                drag.current = { start: [row, column], end: [row, column] };
                setPaintRange(drag.current);
                host.current?.focus();
                host.current?.setPointerCapture(event.pointerId);
              }
            }}
            onClick={() => {
              if (suppressClick.current) {
                suppressClick.current = false;
                return;
              }
              if (!busy) {
                if (onPaint) {
                  onPaint([key]);
                } else {
                  onSelect(key);
                }
                host.current?.focus();
              }
            }} onDoubleClick={() => {
              if (!busy && canEditValue) {
                onSelect(key);
                setEditing(cellEditValue(cell));
              }
            }}>
            {selected === key && editing !== null ? <input aria-label={'Edit ' + key} autoFocus value={editing} onChange={event => setEditing(event.target.value)} onClick={event => event.stopPropagation()} onKeyDown={event => {
              event.stopPropagation();
              if (event.key === 'Enter') {
                event.preventDefault();
                onEdit(editing);
                setEditing(null);
                host.current?.focus();
              } else if (event.key === 'Escape') {
                setEditing(null);
                host.current?.focus();
              }
            }} /> : cell?.formula != null && !(cell.hasCachedValue ?? !!cell.value) ? <span className="formula-pending">={cell.formula}</span> : <span className="cell-text" style={textAppearance(style)}>{displayValue(cell, style, date1904)}</span>}
          </div>;
        })}
      </div>)}
      {(sheet.images ?? []).map(image => <WorksheetImage key={image.id} image={image} selected={selectedImage === image.id} busy={busy || !!onPaint}
        offsetX={48 - columnPosition(sheet, firstColumn)} offsetY={32 - rowPosition(sheet, base + 1)}
        onSelect={() => onSelectImage?.(image.id)} onSave={bounds => onImageEdit?.(image.id, 'transform-image', bounds)} onDelete={() => onImageEdit?.(image.id, 'delete-image')} />)}
    </div>
  </div>;
}
