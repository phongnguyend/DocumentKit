import { test, expect } from '@playwright/test';
import { textOverflowColumns } from '../src/cellOverflow';
import { defaultStyle, type Cell } from '../src/types';
import { cellAlignment, textDirection } from '../src/formatting';
import { execFileSync } from 'node:child_process';
import { resolve } from 'node:path';

test('text overflow respects alignment, occupied cells, merges, wrapping and hidden columns', () => {
  const cell: Cell = { address: 'C1', type: 'inlineStr', value: 'A long label', formula: null, style: 0 };
  const cells = new Map([['C1', cell], ['F1', { ...cell, address: 'F1', value: 'Stop' }]]);
  const columns = [1, 2, 3, 4, 5, 6];
  const bounds = (style = defaultStyle) => textOverflowColumns(cell, style, 1, 2, columns, cells, new Set(), []);
  expect(bounds()).toEqual([2, 4]);
  expect(bounds({ ...defaultStyle, alignment: 'right' })).toEqual([0, 2]);
  expect(bounds({ ...defaultStyle, alignment: 'general', readingOrder: 2 })).toEqual([0, 2]);
  expect(bounds({ ...defaultStyle, alignment: 'left', readingOrder: 2 })).toEqual([2, 4]);
  const hebrew = { ...cell, value: 'שלום עולם' };
  expect(cellAlignment({ ...defaultStyle, alignment: 'general' }, hebrew)).toBe('right');
  expect(textDirection({ ...defaultStyle, readingOrder: 1 }, hebrew)).toBe('ltr');
  expect(bounds({ ...defaultStyle, alignment: 'center' })).toEqual([0, 4]);
  expect(bounds({ ...defaultStyle, wrapText: true })).toEqual([2, 2]);
  expect(bounds({ ...defaultStyle, textRotation: 90 })).toEqual([2, 2]);
  cells.set('D1', { ...cell, address: 'D1', value: '', formula: '""' });
  expect(bounds()).toEqual([2, 2]);
  expect(textOverflowColumns(cell, defaultStyle, 1, 2, columns, cells, new Set([4]), [])).toEqual([2, 4]);
  cells.delete('D1');
  expect(textOverflowColumns(cell, defaultStyle, 1, 2, columns, cells, new Set(), [{ first: [1, 4], last: [2, 5] }])).toEqual([2, 2]);
  expect(textOverflowColumns({ ...cell, type: 'n', value: '12345' }, defaultStyle, 1, 2, columns, cells, new Set(), [])).toEqual([2, 2]);
});

test('right aligned and RTL general text overflow left without reversing Latin text', async ({ page }, info) => {
  const fixture = info.outputPath('direction.xlsx');
  execFileSync('dotnet', ['run', '--project', resolve('../XlsxObjectKit.Tests'), '-c', 'Release', '--no-build', '--', '--write-formatting-fixture', fixture]);
  await page.goto('/');
  await page.getByLabel('Open XLSX', { exact: true }).setInputFiles(fixture);
  for (const key of ['E6', 'E7']) {
    const cell = page.getByRole('gridcell', { name: key, exact: true });
    const overflow = page.locator(`[data-overflow-cell="${key}"]`);
    await expect(overflow).toHaveCSS('text-align', 'right');
    await expect(overflow.locator('span')).toHaveCSS('direction', key === 'E6' ? 'rtl' : 'ltr');
    const box = (await cell.boundingBox())!;
    const ink = await overflow.locator('span').evaluate(span => {
      const range = document.createRange();
      range.selectNodeContents(span);
      const rect = range.getBoundingClientRect();
      return { left: rect.left, right: rect.right };
    });
    expect(ink.left).toBeLessThan(box.x);
    expect(ink.right).toBeLessThanOrEqual(box.x + box.width);
  }
  await page.screenshot({ path: info.outputPath('right-to-left.png') });
});

test('overflow stays visible across empty cells without changing selection or editing', async ({ page, request }, info) => {
  const created = await request.post('/api/documents/new');
  const prepared = await request.post('/api/documents/edit', { multipart: {
    file: { name: 'overflow.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await created.body() },
    edit: JSON.stringify({ action: 'cells', sheetId: '1', cells: [
      { address: 'B1', value: 'Check list for the entire review', type: 'text' },
      { address: 'F1', value: 'Occupied', type: 'text' }
    ] })
  } });
  expect(prepared.ok()).toBeTruthy();
  await page.goto('/');
  await page.getByLabel('Open XLSX', { exact: true }).setInputFiles({ name: 'overflow.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await prepared.body() });
  const cell = (name: string) => page.getByRole('gridcell', { name, exact: true });
  const overflow = page.locator('[data-overflow-cell="B1"]');
  await expect(overflow).toBeVisible();
  await expect(overflow).toHaveCSS('pointer-events', 'none');
  const source = (await cell('B1').boundingBox())!;
  const stop = (await cell('F1').boundingBox())!;
  const painted = (await overflow.boundingBox())!;
  expect(painted.width).toBeGreaterThan(source.width * 3);
  expect(painted.x + painted.width).toBeLessThanOrEqual(stop.x);
  const textWidth = await overflow.locator('span').evaluate(span => {
    const range = document.createRange();
    range.selectNodeContents(span);
    return range.getBoundingClientRect().width;
  });
  expect(textWidth).toBeGreaterThan(source.width);
  await cell('C1').click();
  await expect(cell('C1')).toHaveAttribute('aria-selected', 'true');
  await page.screenshot({ path: info.outputPath('overflow.png') });
  await cell('C1').dblclick();
  await expect(overflow).toHaveCount(0);
  await page.getByRole('textbox', { name: 'Edit C1', exact: true }).fill('Now occupied');
  await page.keyboard.press('Enter');
  await expect(cell('C1')).toHaveText('Now occupied');
  await expect(overflow).toHaveCount(0);
  await expect(cell('B1').locator('.cell-text')).toHaveCSS('opacity', '1');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(overflow).toBeVisible();
});
