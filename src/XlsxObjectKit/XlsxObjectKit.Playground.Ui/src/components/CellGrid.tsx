import { useEffect, useMemo, useRef, useState } from 'react';
import { cellAppearance, displayValue, textAppearance } from '../formatting';
import { address, cellEditValue, defaultStyle, parse, type Sheet, type Style } from '../types';

export function CellGrid({ sheet, styles, selected, onSelect, onEdit, onPaste, busy, date1904, canEditValue = true }: {
  sheet: Sheet; styles: Style[]; selected: string; onSelect: (value: string) => void;
  onEdit: (value: string) => void; onPaste: (value: string) => void; busy: boolean; date1904?: boolean; canEditValue?: boolean;
}) {
  const [scroll, setScroll] = useState(0);
  const [viewportHeight, setViewportHeight] = useState(700);
  const [editing, setEditing] = useState<string | null>(null);
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

  return <div className="cell-grid" ref={host} role="grid" aria-label={sheet.name + ' cells'} aria-rowcount={1048576} aria-colcount={16384} tabIndex={0}
    onScroll={event => setScroll(event.currentTarget.scrollTop)} onCopy={event => {
      if (editing !== null) {
        return;
      }
      event.clipboardData.setData('text/plain', cells.get(selected)?.value ?? '');
      event.preventDefault();
    }} onPaste={event => {
      if (busy || editing !== null) {
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
      } else if (canEditValue && (event.key === 'Enter' || event.key === 'F2')) {
        event.preventDefault();
        const cell = cells.get(selected);
        setEditing(cellEditValue(cell));
      } else if (event.key === 'Delete') {
        event.preventDefault();
        onEdit('');
      }
    }}>
    <div className="grid-surface" style={{ width: x.at(-1), height: y.at(-1) }}>
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
            className={'grid-cell' + (selected === key ? ' selected' : '')} style={{ top: y[index], left: x[ci], width, height, ...cellAppearance(style, cell) }}
            onClick={() => {
              if (!busy) {
                onSelect(key);
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
    </div>
  </div>;
}
