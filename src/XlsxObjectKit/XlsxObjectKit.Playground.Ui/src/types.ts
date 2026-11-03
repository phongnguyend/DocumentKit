export type Cell = { address: string; value: string; type: string; formula: string | null; hasCachedValue?: boolean; style: number };
export type Style = {
  bold: boolean; italic: boolean; color: string; background: string; alignment: string; numberFormat?: string;
  numberFormatId?: number; fontName?: string; fontSize?: number; underline?: string; strike?: boolean;
  wrapText?: boolean; verticalAlignment?: string; indent?: number; textRotation?: number;
  borders?: Record<string, { style: string; color: string }>;
};
export type SheetImage = { id: string; name: string; dataUrl: string; x: number; y: number; width: number; height: number; editable: boolean };
export type Sheet = { images?: SheetImage[]; id: string; name: string; hidden: boolean; cells: Cell[]; merges: string[]; rowHeights: Record<number, number>; columnWidths: Record<number, number>; rowStyles?: Record<number, number>; columnStyles?: Record<number, number>; defaultRowHeight?: number; defaultColumnWidth?: number };
export type Inspection = { sheets: Sheet[]; styles: Style[]; warnings: string[]; date1904?: boolean };
export type Snapshot = { bytes: Blob; info: Inspection };
export const defaultStyle: Style = { bold: false, italic: false, color: '#24334a', background: '#ffffff', alignment: 'left' };
export function cellValueType(cell?: Cell): string {
  if (cell?.formula != null) {
    return 'formula';
  }
  if (!cell || (cell.type === 'n' && cell.value === '')) {
    return 'auto';
  }
  return ({ n: 'number', b: 'boolean', s: 'text', inlineStr: 'text', str: 'text', d: 'date', e: 'error' } as Record<string, string>)[cell.type] ?? 'auto';
}
export function cellEditValue(cell?: Cell): string {
  if (cell?.formula != null) {
    return '=' + cell.formula;
  }
  return cell?.type === 'b' ? cell.value === '1' ? 'true' : 'false' : cell?.value ?? '';
}
export function address(row: number, column: number) {
  let text = '';
  while (column > 0) {
    column--;
    text = String.fromCharCode(65 + column % 26) + text;
    column = Math.floor(column / 26);
  }
  return text + row;
}
export function parse(value: string): [number, number] {
  const match = /^([A-Z]{1,3})([1-9]\d{0,6})$/i.exec(value);
  if (!match) {
    throw new Error('Enter an A1 cell address.');
  }
  let column = 0;
  for (const letter of match[1].toUpperCase()) {
    column = column * 26 + letter.charCodeAt(0) - 64;
  }
  if (+match[2] > 1048576 || column > 16384) {
    throw new Error('Cell is outside worksheet limits.');
  }
  return [+match[2], column];
}
export async function post(path: string, bytes?: Blob, edit?: object) {
  let body: FormData | Blob | undefined = bytes;
  if (edit && bytes) {
    const form = new FormData();
    form.append('file', bytes, 'workbook.xlsx');
    form.append('edit', JSON.stringify(edit));
    body = form;
  }
  const response = await fetch('/api/documents/' + path, { method: 'POST', body });
  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    throw new Error(error.detail ?? 'Request failed.');
  }
  return response;
}
export async function inspect(bytes: Blob): Promise<Snapshot> {
  return { bytes, info: await (await post('inspect', bytes)).json() };
}

export function columnPosition(sheet: Sheet, column: number) {
  const width = sheet.defaultColumnWidth ?? 8.43;
  return (column - 1) * (width * 7 + 5) + Object.entries(sheet.columnWidths).filter(([key]) => +key < column).reduce((sum, [, value]) => sum + (value - width) * 7, 0);
}
export function rowPosition(sheet: Sheet, row: number) {
  const height = sheet.defaultRowHeight ?? 15;
  return ((row - 1) * height + Object.entries(sheet.rowHeights).filter(([key]) => +key < row).reduce((sum, [, value]) => sum + value - height, 0)) * 4 / 3;
}
