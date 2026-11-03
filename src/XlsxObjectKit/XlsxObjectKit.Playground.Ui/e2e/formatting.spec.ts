import { test, expect } from '@playwright/test';
import { execFileSync } from 'node:child_process';
import { resolve } from 'node:path';
import { displayValue } from '../src/formatting';
import { defaultStyle, type Cell } from '../src/types';

test('spreadsheet formats preserve text and handle numeric and date values', () => {
  function show(value: string, numberFormat: string, date1904 = false, type = 'n') {
    const cell: Cell = { address: 'A1', value, type, style: 0, formula: null };
    return displayValue(cell, { ...defaultStyle, numberFormat }, date1904);
  }
  expect(show('1234.5', '#,##0.00')).toBe('1,234.50');
  expect(show('0.125', '0.00%')).toBe('12.50%');
  expect(show('-1234.5', '"$"#,##0.00;("$"#,##0.00)')).toBe('($1,234.50)');
  expect(show('0', 'yyyy-mm-dd', true)).toBe('1904-01-01');
  expect(show('61', 'yyyy-mm-dd')).toBe('1900-03-01');
  expect(show('60', 'yyyy-mm-dd')).toBe('1900-02-29');
  expect(show('1.5', '[h]:mm')).toBe('36:00');
  expect(show('1234', '0.00E+00')).toBe('1.23E+03');
  expect(show('00123', '#,##0.00', false, 'inlineStr')).toBe('00123');
  expect(show('0', ';;;')).toBe('');
  expect(show('12', '"unterminated')).toBe('12');
});

test('uploaded workbook displays its styles and formats without changing raw values', async ({ page }, info) => {
  const fixture = info.outputPath('formatting.xlsx');
  execFileSync('dotnet', ['run', '--project', resolve('../XlsxObjectKit.Tests'), '-c', 'Release', '--no-build', '--', '--write-formatting-fixture', fixture]);
  await page.goto('/');
  await page.getByLabel('Open XLSX', { exact: true }).setInputFiles(fixture);
  const a1 = page.getByRole('gridcell', { name: 'A1', exact: true });
  await expect(a1).toHaveText('1,234.50');
  await expect(a1).toHaveCSS('font-weight', '700');
  await expect(a1).toHaveCSS('font-style', 'italic');
  await expect(a1).toHaveCSS('color', 'rgb(0, 128, 0)');
  await expect(a1).toHaveCSS('background-color', 'rgb(64, 255, 64)');
  await expect(a1).toHaveCSS('border-left-color', 'rgb(0, 0, 255)');
  await expect(a1).toHaveCSS('border-bottom-style', 'double');
  await expect(a1).toHaveCSS('text-decoration-line', 'underline line-through');
  await expect(a1).toHaveCSS('justify-content', 'center');
  await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveValue('1234.5');
  await expect(page.getByRole('gridcell', { name: 'A2', exact: true })).toHaveText('12.50%');
  await expect(page.getByRole('gridcell', { name: 'A3', exact: true })).toHaveText('1904-01-01');
  await expect(page.getByRole('gridcell', { name: 'A4', exact: true }).locator('.cell-text')).toHaveCSS('white-space', 'pre-wrap');
  await expect(page.getByRole('gridcell', { name: 'B5', exact: true })).toHaveCSS('font-weight', '700');
  await expect(page.getByRole('gridcell', { name: 'B4', exact: true })).toHaveCSS('font-weight', '400');
  await page.screenshot({ path: info.outputPath('formatting.png') });
  await a1.dblclick();
  await expect(page.getByRole('textbox', { name: 'Edit A1' })).toHaveValue('1234.5');
  await page.getByRole('textbox', { name: 'Edit A1' }).fill('2500');
  await page.getByRole('textbox', { name: 'Edit A1' }).press('Enter');
  await expect(a1).toHaveText('2,500.00');
});
