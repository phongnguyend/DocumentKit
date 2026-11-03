import SSF from 'ssf';
import type { CSSProperties } from 'react';
import type { Cell, Style } from './types';

export function displayValue(cell: Cell | undefined, style: Style, date1904 = false): string {
  if (!cell) {
    return '';
  }
  if (cell.type === 'b') {
    return cell.value === '1' ? 'TRUE' : 'FALSE';
  }
  if (cell.value === '' || !['n', 's', 'inlineStr', 'str'].includes(cell.type)) {
    return cell.value;
  }
  const value = cell.type === 'n' ? Number(cell.value) : cell.value;
  if (typeof value === 'number' && !Number.isFinite(value)) {
    return cell.value;
  }
  try {
    return SSF.format(style.numberFormat || style.numberFormatId || 'General', value, { date1904 });
  } catch {
    return cell.value;
  }
}

export function textDirection(style: Style, cell?: Cell): 'rtl' | 'ltr' {
  if (style.readingOrder === 2) {
    return 'rtl';
  }
  if (style.readingOrder === 1) {
    return 'ltr';
  }
  const firstLetter = cell?.value.match(/\p{Letter}/u)?.[0] ?? '';
  return /[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}\p{Script=Adlam}]/u.test(firstLetter) ? 'rtl' : 'ltr';
}

export function cellAlignment(style: Style, cell?: Cell): 'left' | 'right' | 'center' | 'justify' {
  const numeric = cell?.type === 'n' && cell.value !== '';
  return style.alignment === 'general' ? numeric ? 'right' : cell?.type === 'b' || cell?.type === 'e' ? 'center' : textDirection(style, cell) === 'rtl' ? 'right' : 'left'
    : ['center', 'centerContinuous'].includes(style.alignment) ? 'center'
      : ['justify', 'distributed'].includes(style.alignment) ? 'justify' : style.alignment === 'right' ? 'right' : 'left';
}

export function cellAppearance(style: Style, cell?: Cell): CSSProperties {
  const alignment = cellAlignment(style, cell);
  const result: CSSProperties = {
    color: style.color, background: style.background,
    fontWeight: style.bold ? 700 : 400, fontStyle: style.italic ? 'italic' : 'normal',
    fontFamily: style.fontName ? `${JSON.stringify(style.fontName)}, Calibri, Arial, sans-serif` : undefined,
    fontSize: style.fontSize ? `${style.fontSize * 4 / 3}px` : undefined,
    textAlign: alignment,
    justifyContent: style.verticalAlignment === 'top' ? 'flex-start' : style.verticalAlignment === 'center' ? 'center' : 'flex-end',
    paddingLeft: alignment !== 'right' ? 4 + (style.indent ?? 0) * 12 : 4,
    paddingRight: alignment === 'right' ? 4 + (style.indent ?? 0) * 12 : 4,
    textDecorationLine: [style.underline && style.underline !== 'none' ? 'underline' : '', style.strike ? 'line-through' : ''].filter(Boolean).join(' ') || 'none',
    textDecorationStyle: style.underline?.startsWith('double') ? 'double' : 'solid'
  };
  for (const [side, edge] of Object.entries(style.borders ?? {})) {
    const key = ({ left: 'borderLeft', right: 'borderRight', top: 'borderTop', bottom: 'borderBottom' } as const)[side as 'left'];
    if (!key) {
      continue;
    }
    const width = edge.style === 'thick' || edge.style === 'double' ? 3 : edge.style.startsWith('medium') ? 2 : 1;
    const line = edge.style === 'double' ? 'double' : edge.style === 'dotted' || edge.style === 'hair' ? 'dotted' : /dash/i.test(edge.style) ? 'dashed' : 'solid';
    result[key] = `${width}px ${line} ${edge.color}`;
  }
  return result;
}

export function textAppearance(style: Style, cell?: Cell): CSSProperties {
  const rotation = style.textRotation ?? 0;
  return {
    direction: textDirection(style, cell),
    unicodeBidi: 'isolate',
    whiteSpace: style.wrapText ? 'pre-wrap' : 'pre',
    overflowWrap: style.wrapText ? 'anywhere' : undefined,
    writingMode: rotation === 255 ? 'vertical-rl' : undefined,
    textOrientation: rotation === 255 ? 'upright' : undefined,
    transform: rotation > 0 && rotation <= 180 ? `rotate(${rotation <= 90 ? -rotation : rotation - 90}deg)` : undefined
  };
}
