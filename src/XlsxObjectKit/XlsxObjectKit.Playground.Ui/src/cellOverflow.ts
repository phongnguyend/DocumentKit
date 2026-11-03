import { address, type Cell, type Style } from './types';
import { cellAlignment } from './formatting';

type Merge = { first: [number, number]; last: [number, number] };

// Return column-window indexes; never enlarge the actual cell's hit target.
export function textOverflowColumns(cell: Cell | undefined, style: Style, row: number, index: number,
  columns: number[], cells: Map<string, Cell>, hidden: Set<number>, merges: Merge[], editingCell?: string, limits: [number, number] = [0, columns.length - 1]): [number, number] {
  const contains = (column: number) => merges.some(m => row >= m.first[0] && row <= m.last[0] && column >= m.first[1] && column <= m.last[1]);
  if (!cell || !['s', 'str', 'inlineStr'].includes(cell.type) || !cell.value || style.wrapText || style.textRotation
    || (cell.formula != null && !(cell.hasCachedValue ?? !!cell.value)) || contains(columns[index])) {
    return [index, index];
  }
  const alignment = cellAlignment(style, cell);
  if (!['general', 'left', 'right', 'center'].includes(style.alignment)) {
    return [index, index];
  }
  function edge(direction: number) {
    let end = index;
    for (let next = index + direction; next >= limits[0] && next <= limits[1]; next += direction) {
      if (hidden.has(columns[next])) {
        continue;
      }
      const neighbor = cells.get(address(row, columns[next]));
      // An empty formula result still occupies its cell, unlike a styled blank.
      if (address(row, columns[next]) === editingCell || contains(columns[next]) || neighbor?.formula != null || (neighbor && neighbor.value !== '')) {
        break;
      }
      end = next;
    }
    return end;
  }
  return [alignment === 'left' ? index : edge(-1), alignment === 'right' ? index : edge(1)];
}
