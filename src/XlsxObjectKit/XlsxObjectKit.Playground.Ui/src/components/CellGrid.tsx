import { conditionalStyle } from '../conditionalFormatting';
import { useEffect, useMemo, useRef, useState, type PointerEvent, type ReactNode } from 'react';
import { textOverflowColumns } from '../cellOverflow';
import { DimensionHandle } from './DimensionHandle';
import { WorksheetImage } from './WorksheetImage';
import { columnPosition, rowPosition } from '../types';
import { cellAppearance, displayValue, textAppearance } from '../formatting';
import { address, cellEditValue, defaultStyle, parse, type Sheet, type Style } from '../types';

export function CellGrid({ sheet, styles, selected, onSelect, onEdit, onPaste, busy, date1904, canEditValue = true, onPaint, paintSource, selectedImage, onSelectImage, onImageEdit, onClearRange, onRangeChange, onSelectionAxisChange, onResize }: {
  onResize?: (axis: 'row' | 'column', cell: string, size: number) => void;
  onRangeChange?: (range: string) => void;
  onSelectionAxisChange?: (axis: 'row' | 'column' | null) => void;
  onClearRange?: (addresses: string[]) => void;
  selectedImage?: string; onSelectImage?: (id: string) => void; onImageEdit?: (id: string, action: string, bounds?: object) => void;
  sheet: Sheet; styles: Style[]; selected: string; onSelect: (value: string) => void;
  onEdit: (value: string) => void; onPaste: (value: string) => void; busy: boolean; date1904?: boolean; canEditValue?: boolean; onPaint?: (addresses: string[]) => void; paintSource?: string;
}) {
  const [scroll, setScroll] = useState(0);
  const [scrollLeft, setScrollLeft] = useState(0);
  const [viewportWidth, setViewportWidth] = useState(1000);
  const [viewportHeight, setViewportHeight] = useState(700);
  const [editing, setEditing] = useState<string | null>(null);
  const [paintRange, setPaintRange] = useState<{ start: [number, number]; end: [number, number]; axis?: 'row' | 'column'; anchor?: string } | null>(null);
  const [selection, setSelection] = useState<typeof paintRange>(null);
  const drag = useRef<typeof paintRange>(null);
  const suppressClick = useRef(false);
  const preserveScroll = useRef(false);
  const host = useRef<HTMLDivElement>(null);
  const [selectedRow, selectedColumn] = parse(selected);
  // Bound the browser scroll area; Go to cell opens any other worksheet window.
  const base = Math.floor((selectedRow - 1) / 10000) * 10000;
  const firstColumn = 1;
  const usedLastColumn = useMemo(() => {
    let last = 26;
    for (const cell of sheet.cells) {
      if (cell.value !== '' || cell.formula != null) {
        last = Math.max(last, parse(cell.address)[1]);
      }
    }
    for (const range of sheet.merges) {
      last = Math.max(last, parse(range.split(':').at(-1)!)[1]);
    }
    const imageRight = Math.max(0, ...(sheet.images ?? []).map(image => image.x + image.width));
    const hidden = new Set(sheet.hiddenColumns ?? []);
    let offset = 0;
    for (let column = 1; column <= 16384 && offset < imageRight; column++) {
      offset += hidden.has(column) ? 0 : Math.max(1, Math.min(255, sheet.columnWidths[column] ?? sheet.defaultColumnWidth ?? 8.43)) * 7 + 5;
      last = Math.max(last, column);
    }
    return Math.min(16384, last);
  }, [sheet]);
  const columnCount = Math.max(usedLastColumn, selectedColumn);
  const columns = useMemo(() => Array.from({ length: columnCount }, (_, i) => i + 1), [columnCount]);
  const count = Math.min(10000, 1048576 - base);
  const conditionalRules = useMemo(() => [...(sheet.conditionalRules ?? [])].sort((a, b) => a.priority - b.priority), [sheet]);
  const cells = useMemo(() => new Map(sheet.cells.map(cell => [cell.address, cell])), [sheet]);
  const hiddenRows = useMemo(() => new Set(sheet.hiddenRows ?? []), [sheet]);
  const hiddenColumns = useMemo(() => new Set(sheet.hiddenColumns ?? []), [sheet]);
  const widths = useMemo(() => columns.map(column => hiddenColumns.has(column) ? 0 : Math.max(1, Math.min(255, sheet.columnWidths[column] ?? sheet.defaultColumnWidth ?? 8.43)) * 7 + 5), [sheet, hiddenColumns, columns]);
  const x = useMemo(() => {
    const offsets = [48];
    widths.forEach(width => offsets.push(offsets[offsets.length - 1] + width));
    return offsets;
  }, [widths]);
  const firstVisibleColumn = Math.max(0, x.findIndex((offset, index) => index < columns.length && x[index + 1] > scrollLeft) - 3);
  const rightEdge = x.findIndex(offset => offset >= scrollLeft + viewportWidth);
  const lastVisibleColumn = Math.min(columns.length, (rightEdge < 0 ? columns.length : rightEdge) + 3);
  const visibleColumns = columns.slice(firstVisibleColumn, lastVisibleColumn);
  const storedColumns = useMemo(() => {
    const result = new Map<number, number[]>();
    for (const cell of sheet.cells) {
      if (cell.value !== '' || cell.formula != null) {
        const [row, column] = parse(cell.address);
        const list = result.get(row) ?? [];
        list.push(column);
        result.set(row, list);
      }
    }
    for (const list of result.values()) {
      list.sort((a, b) => a - b);
    }
    return result;
  }, [sheet]);
  const y = useMemo(() => {
    const offsets = [32];
    for (let i = 1; i <= count; i++) {
      offsets.push(offsets[i - 1] + (hiddenRows.has(base + i) ? 0 : Math.max(1, Math.min(409, sheet.rowHeights[base + i] ?? sheet.defaultRowHeight ?? 15)) * 4 / 3));
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
      setViewportWidth(entries[0]?.contentRect.width ?? 1000);
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
    if (preserveScroll.current) {
      preserveScroll.current = false;
      return;
    }
    const left = x[selectedColumn - 1];
    const right = x[selectedColumn];
    if (host.current && (left < host.current.scrollLeft || right > host.current.scrollLeft + host.current.clientWidth)) {
      host.current.scrollLeft = left < host.current.scrollLeft ? left : right - host.current.clientWidth;
      setScrollLeft(host.current.scrollLeft);
    }
    const top = y[selectedRow - base - 1];
    if (host.current && (top < host.current.scrollTop + 32 || top > host.current.scrollTop + host.current.clientHeight - 40)) {
      host.current.scrollTop = Math.max(0, top - 50);
    }
  }, [selected, sheet.id]);

  useEffect(() => {
    setSelection(current => current && (current.anchor ?? address(...current.start)) === selected && !selectedImage ? current : null);
  }, [selected, selectedImage]);
  useEffect(() => {
    setSelection(null);
  }, [sheet.id, !!onPaint]);
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
        setSelection(null);
      }
    }
    window.addEventListener('keydown', escape);
    return () => window.removeEventListener('keydown', escape);
  }, []);
  function selectHeader(axis: 'row' | 'column', index: number) {
    if (busy || onPaint) {
      return;
    }
    cancelDrag();
    setEditing(null);
    const anchor = axis === 'column' ? address(selectedRow, index) : address(index, selectedColumn);
    preserveScroll.current = anchor !== selected;
    setSelection({
      start: axis === 'column' ? [1, index] : [index, 1],
      end: axis === 'column' ? [1048576, index] : [index, 16384],
      axis,
      anchor,
    });
    onSelect(anchor);
    host.current?.focus();
  }
  function startHeaderDrag(event: PointerEvent<HTMLButtonElement>, axis: 'row' | 'column', index: number) {
    if (busy || onPaint || event.button !== 0) {
      return;
    }
    event.preventDefault();
    setEditing(null);
    drag.current = {
      start: axis === 'column' ? [1, index] : [index, 1],
      end: axis === 'column' ? [1048576, index] : [index, 16384],
      axis,
      anchor: axis === 'column' ? address(selectedRow, index) : address(index, selectedColumn),
    };
    setPaintRange(drag.current);
    host.current?.focus({ preventScroll: true });
    host.current?.setPointerCapture(event.pointerId);
  }
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
    drag.current = { ...drag.current, end: drag.current.axis === 'column' ? [1048576, column] : drag.current.axis === 'row' ? [base + ri + 1, 16384] : [base + ri + 1, column] };
    setPaintRange(drag.current);
  }
  function boundsOf(range: NonNullable<typeof paintRange>) {
    let top = Math.min(range.start[0], range.end[0]);
    let bottom = Math.max(range.start[0], range.end[0]);
    let left = Math.min(range.start[1], range.end[1]);
    let right = Math.max(range.start[1], range.end[1]);
    // Header selections describe an exact worksheet axis, even across merged cells.
    if (range.axis) {
      return { top, bottom, left, right };
    }
    let changed = true;
    while (changed) {
      changed = false;
      for (const merge of merges) {
        if (merge.first[0] <= bottom && merge.last[0] >= top && merge.first[1] <= right && merge.last[1] >= left) {
          const next = [Math.min(top, merge.first[0]), Math.max(bottom, merge.last[0]), Math.min(left, merge.first[1]), Math.max(right, merge.last[1])];
          if (next[0] !== top || next[1] !== bottom || next[2] !== left || next[3] !== right) {
            [top, bottom, left, right] = next;
            changed = true;
          }
        }
      }
    }
    return { top, bottom, left, right };
  }
  const visibleRange = paintRange ?? selection;
  const { top: rangeTop, bottom: rangeBottom, left: rangeLeft, right: rangeRight } = visibleRange ? boundsOf(visibleRange) : { top: 0, bottom: 0, left: 0, right: 0 };

  useEffect(() => {
    onRangeChange?.(visibleRange ? address(rangeTop, rangeLeft) + ':' + address(rangeBottom, rangeRight) : selected);
  }, [rangeTop, rangeLeft, rangeBottom, rangeRight, selected, !!visibleRange, onRangeChange]);

  useEffect(() => {
    onSelectionAxisChange?.(visibleRange?.axis ?? null);
  }, [visibleRange?.axis, onSelectionAxisChange]);

  function rowColumns(row: number) {
    const result = new Set(visibleColumns);
    // Keep offscreen anchors whose merged area or overflowing text can still be visible.
    for (const merge of merges) {
      if (row >= merge.first[0] && row <= merge.last[0] && merge.last[1] >= firstVisibleColumn + 1 && merge.first[1] <= lastVisibleColumn) {
        let anchor = merge.first[1];
        while (hiddenColumns.has(anchor) && anchor <= merge.last[1]) {
          anchor++;
        }
        result.add(anchor);
      }
    }
    let before: number | undefined;
    for (const column of storedColumns.get(row) ?? []) {
      if (hiddenColumns.has(column)) {
        continue;
      }
      if (column < firstVisibleColumn + 1) {
        before = column;
      } else if (column > lastVisibleColumn) {
        result.add(column);
        break;
      }
    }
    if (before !== undefined) {
      result.add(before);
    }
    return [...result].sort((a, b) => a - b);
  }
  const overflowText: ReactNode[] = [];
  return <div className={'cell-grid' + (onPaint ? ' painting' : '')} ref={host} role="grid" aria-label={sheet.name + ' cells'} aria-rowcount={1048576} aria-colcount={16384} tabIndex={0}
    onPointerMove={event => updateDrag(event.clientX, event.clientY)}
    onPointerCancel={cancelDrag} onLostPointerCapture={cancelDrag}
    onPointerUp={event => {
      if (!drag.current || busy) {
        return;
      }
      updateDrag(event.clientX, event.clientY);
      const range = drag.current;
      if (!onPaint) {
        cancelDrag();
        setSelection(range);
        const anchor = range.anchor ?? address(...range.start);
        if (range.axis) {
          preserveScroll.current = anchor !== selected;
        }
        onSelect(anchor);
        if (event.currentTarget.hasPointerCapture(event.pointerId)) {
          event.currentTarget.releasePointerCapture(event.pointerId);
        }
        return;
      }
      const targets: string[] = [];
      const expanded = boundsOf(range);
      for (let row = expanded.top; row <= expanded.bottom && targets.length <= 10000; row++) {
        for (let column = expanded.left; column <= expanded.right && targets.length <= 10000; column++) {
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
    onScroll={event => {
      setScroll(event.currentTarget.scrollTop);
      setScrollLeft(event.currentTarget.scrollLeft);
    }} onCopy={event => {
      if (editing !== null) {
        return;
      }
      if (selection && (rangeBottom - rangeTop + 1) * (rangeRight - rangeLeft + 1) <= 10000) {
        const rows: string[] = [];
        for (let row = rangeTop; row <= rangeBottom; row++) {
          const values: string[] = [];
          for (let column = rangeLeft; column <= rangeRight; column++) {
            const value = cells.get(address(row, column))?.value ?? '';
            values.push(/[\t\n\r"]/.test(value) ? '"' + value.replaceAll('"', '""') + '"' : value);
          }
          rows.push(values.join('\t'));
        }
        event.clipboardData.setData('text/plain', rows.join('\n'));
      } else if (!selection) {
        event.clipboardData.setData('text/plain', cells.get(selected)?.value ?? '');
      }
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
        setSelection(null);
        event.preventDefault();
        let nextRow = selectedRow + delta[0];
        let nextColumn = selectedColumn + delta[1];
        while (delta[0] && hiddenRows.has(nextRow) && nextRow >= 1 && nextRow <= 1048576) {
          nextRow += delta[0];
        }
        while (delta[1] && hiddenColumns.has(nextColumn) && nextColumn >= 1 && nextColumn <= 16384) {
          nextColumn += delta[1];
        }
        if (nextRow >= 1 && nextRow <= 1048576 && nextColumn >= 1 && nextColumn <= 16384) {
          onSelect(address(nextRow, nextColumn));
        }
      } else if (onPaint && event.key === 'Enter') {
        event.preventDefault();
        onPaint([selected]);
      } else if (canEditValue && (event.key === 'Enter' || event.key === 'F2')) {
        event.preventDefault();
        const cell = cells.get(selected);
        setEditing(cellEditValue(cell));
      } else if (!onPaint && event.key === 'Delete') {
        event.preventDefault();
        if (selection && onClearRange) {
          const targets: string[] = [];
          if (selection.axis) {
            for (const cell of sheet.cells) {
              const [row, column] = parse(cell.address);
              if (row >= rangeTop && row <= rangeBottom && column >= rangeLeft && column <= rangeRight && (cell.value !== '' || cell.formula != null)) {
                targets.push(cell.address);
                if (targets.length > 10000) {
                  break;
                }
              }
            }
          } else {
            for (let row = rangeTop; row <= rangeBottom && targets.length <= 10000; row++) {
              for (let column = rangeLeft; column <= rangeRight && targets.length <= 10000; column++) {
                const merge = merges.find(item => row >= item.first[0] && row <= item.last[0] && column >= item.first[1] && column <= item.last[1]);
                if (!merge || (row === merge.first[0] && column === merge.first[1])) {
                  targets.push(address(row, column));
                }
              }
            }
          }
          if (targets.length > 0) {
            onClearRange(targets);
          }
        } else {
          onEdit('');
        }
      }
    }}>
    <div className="grid-surface" style={{ width: x.at(-1), height: y.at(-1) }}>
      {visibleRange && <div className={onPaint ? "paint-range" : "selected-range"} aria-hidden="true" style={{ left: x[Math.max(0, rangeLeft - firstColumn)], top: y[Math.max(0, rangeTop - base - 1)], width: x[Math.min(columns.length, rangeRight - firstColumn + 1)] - x[Math.max(0, rangeLeft - firstColumn)], height: y[Math.min(count, rangeBottom - base)] - y[Math.max(0, rangeTop - base - 1)] }} />}
      <div className="column-headings" role="row" style={{ width: x.at(-1) }}><div className="corner" />{visibleColumns.map(column => hiddenColumns.has(column) ? null : <div role="columnheader" aria-selected={visibleRange?.axis === 'column' && column >= rangeLeft && column <= rangeRight} key={column} style={{ left: x[column - 1], width: widths[column - 1] }}><button type="button" className="header-select" onKeyDown={event => event.stopPropagation()} aria-label={'Select column ' + address(1, column).replace('1', '')} disabled={busy || !!onPaint} onPointerDown={event => startHeaderDrag(event, 'column', column)} onClick={event => {
        if (event.detail === 0) {
          selectHeader('column', column);
        }
      }}>{address(1, column).replace('1', '')}</button><DimensionHandle axis="column" label={address(1, column).replace('1', '')} size={sheet.columnWidths[column] ?? sheet.defaultColumnWidth ?? 8.43} disabled={busy || !!onPaint} onSave={value => onResize?.('column', address(1, column), value)} /></div>)}</div>
      {Array.from({ length: end - start }, (_, i) => i + start).filter(index => !hiddenRows.has(base + index + 1)).map(index => <div key={index} role="row" aria-rowindex={base + index + 1}>
        <div className="row-heading" role="rowheader" aria-selected={visibleRange?.axis === 'row' && base + index + 1 >= rangeTop && base + index + 1 <= rangeBottom} style={{ top: y[index], height: y[index + 1] - y[index] }}><button type="button" className="header-select" onKeyDown={event => event.stopPropagation()} aria-label={'Select row ' + (base + index + 1)} disabled={busy || !!onPaint} onPointerDown={event => startHeaderDrag(event, 'row', base + index + 1)} onClick={event => {
          if (event.detail === 0) {
            selectHeader('row', base + index + 1);
          }
        }}>{base + index + 1}</button><DimensionHandle axis="row" label={String(base + index + 1)} size={sheet.rowHeights[base + index + 1] ?? sheet.defaultRowHeight ?? 15} disabled={busy || !!onPaint} onSave={value => onResize?.('row', address(base + index + 1, 1), value)} /></div>
        {rowColumns(base + index + 1).map(column => {
          const ci = column - 1;
          if (hiddenColumns.has(column)) {
            return null;
          }
          const row = base + index + 1;
          let key = address(row, column);
          const merge = merges.find(item => row >= item.first[0] && row <= item.last[0] && column >= item.first[1] && column <= item.last[1]);
          if (merge) {
            let firstRow = merge.first[0];
            let firstCol = merge.first[1];
            while (hiddenRows.has(firstRow) && firstRow <= merge.last[0]) {
              firstRow++;
            }
            while (hiddenColumns.has(firstCol) && firstCol <= merge.last[1]) {
              firstCol++;
            }
            if (row !== firstRow || column !== firstCol) {
              return null;
            }
            key = address(...merge.first);
          }
          const cell = cells.get(key);
          const style = conditionalStyle(cell, styles[cell?.style ?? sheet.rowStyles?.[row] ?? sheet.columnStyles?.[column] ?? 0] ?? defaultStyle, conditionalRules);
          const width = merge ? x[Math.min(columns.length, merge.last[1] - firstColumn + 1)] - x[ci] : widths[ci];
          const height = merge ? y[Math.min(count, merge.last[0] - base)] - y[index] : y[index + 1] - y[index];
          const appearance = cellAppearance(style, cell);
          const [overflowStart, overflowEnd] = textOverflowColumns(cell, style, row, ci, columns, cells, hiddenColumns, merges, editing !== null ? selected : undefined, [Math.min(ci, firstVisibleColumn), Math.max(ci, lastVisibleColumn - 1)]);
          const overflowing = (overflowStart !== ci || overflowEnd !== ci) && !(selected === key && editing !== null);
          if (overflowing) {
            // Separate paint layer keeps neighboring cells selectable and their backgrounds intact.
            const leftPadding = Number(appearance.paddingLeft ?? 4);
            const rightPadding = Number(appearance.paddingRight ?? 4);
            const contentWidth = Math.max(0, width - leftPadding - rightPadding - 1);
            const anchor = appearance.textAlign === 'right' ? 1 : appearance.textAlign === 'center' ? 0.5 : 0;
            overflowText.push(<div key={key} aria-hidden="true" data-overflow-cell={key} className="cell-overflow"
              style={{ ...appearance, top: y[index], left: x[overflowStart], width: x[overflowEnd + 1] - x[overflowStart] - 1, height: height - 1,
                background: 'transparent', border: 0, padding: '1px 0' }}>
              <span className="cell-text" style={{ ...textAppearance(style, cell), width: 'max-content', transform: `translateX(${-anchor * 100}%)`,
                marginLeft: x[ci] - x[overflowStart] + leftPadding + anchor * contentWidth }}>{displayValue(cell, style, date1904)}</span>
            </div>);
          }
          return <div key={column} role="gridcell" aria-label={key} aria-selected={visibleRange ? row >= rangeTop && row <= rangeBottom && column >= rangeLeft && column <= rangeRight : selected === key} aria-colindex={column}
            className={'grid-cell' + (selected === key && !visibleRange ? ' selected' : '') + (paintSource === key ? ' paint-source' : '')} style={{ top: y[index], left: x[ci], width, height, ...cellAppearance(style, cell) }}
            onPointerDown={event => {
              suppressClick.current = false;
              if (!busy && event.button === 0 && !(event.target instanceof HTMLInputElement)) {
                event.preventDefault();
                suppressClick.current = true;
                drag.current = { start: parse(key), end: [row, column] };
                setPaintRange(drag.current);
                host.current?.focus();
                if (onPaint) {
                  host.current?.setPointerCapture(event.pointerId);
                } else {
                  event.currentTarget.setPointerCapture(event.pointerId);
                }
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
                setSelection(null);
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
            }} /> : cell?.formula != null && !(cell.hasCachedValue ?? !!cell.value) ? <span className="formula-pending">={cell.formula}</span> : <span className="cell-text" style={{ ...textAppearance(style, cell), opacity: overflowing ? 0 : undefined }}>{displayValue(cell, style, date1904)}</span>}
          </div>;
        })}
      </div>)}
      {overflowText}
      {(sheet.images ?? []).map(image => <WorksheetImage key={image.id} image={image} selected={selectedImage === image.id} busy={busy || !!onPaint}
        offsetX={48 - columnPosition(sheet, firstColumn)} offsetY={32 - rowPosition(sheet, base + 1)}
        onSelect={() => onSelectImage?.(image.id)} onSave={bounds => onImageEdit?.(image.id, 'transform-image', bounds)} onDelete={() => onImageEdit?.(image.id, 'delete-image')} />)}
    </div>
    {selection && !onPaint && <span className="range-selection-status" role="status">{address(rangeTop, rangeLeft)}:{address(rangeBottom, rangeRight)} - {(rangeBottom - rangeTop + 1) * (rangeRight - rangeLeft + 1)} cells{(rangeBottom - rangeTop + 1) * (rangeRight - rangeLeft + 1) > 10000 ? (selection.axis ? ' (copy limit: 10,000 cells; clear limit: 10,000 populated cells)' : ' (copy/clear limit: 10,000)') : ''}</span>}
  </div>;
}
