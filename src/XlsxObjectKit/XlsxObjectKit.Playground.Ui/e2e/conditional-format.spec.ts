import { test, expect } from '@playwright/test';
import { conditionalStyle } from '../src/conditionalFormatting';
import { defaultStyle, type ConditionalRule } from '../src/types';
import { execFileSync } from 'node:child_process';
import { resolve } from 'node:path';

test('imported percentage threshold rules render differential backgrounds and theme text', async ({ page }, info) => {
  const fixture = info.outputPath('conditional-reference.xlsx');
  execFileSync('dotnet', ['run', '--project', resolve('../XlsxObjectKit.Tests'), '-c', 'Release', '--no-build', '--', '--write-conditional-fixture', fixture]);
  await page.goto('/');
  await page.getByLabel('Open XLSX', { exact: true }).setInputFiles(fixture);
  const cell = (name: string) => page.getByRole('gridcell', { name, exact: true });
  await expect(cell('A1')).toHaveCSS('background-color', 'rgb(255, 0, 0)');
  await expect(cell('A1')).toHaveCSS('color', 'rgb(255, 255, 255)');
  await expect(cell('A2')).not.toHaveCSS('background-color', 'rgb(255, 0, 0)');
  await cell('D1').dblclick();
  await page.getByRole('textbox', { name: 'Edit D1', exact: true }).fill('300');
  await page.keyboard.press('Enter');
  await expect(cell('A1')).not.toHaveCSS('background-color', 'rgb(255, 0, 0)');
});

test('expression matches apply differential bold and preserve other font properties', () => {
  const rule: ConditionalRule = { id: 0, range: 'A1:A3', operator: 'expression', value: null, secondValue: null,
    color: null, background: null, priority: 1, stopIfTrue: true, supported: true,
    bold: true, italic: false, matchedCells: ['A2'] };
  const cell = { address: 'A2', type: 'str', value: 'Holiday', formula: 'B2', hasCachedValue: true, style: 0 };
  expect(conditionalStyle(cell, { ...defaultStyle, italic: true }, [rule])).toMatchObject({ bold: true, italic: false });
  expect(conditionalStyle({ ...cell, address: 'A1' }, defaultStyle, [rule])).toBe(defaultStyle);
});

test('numeric conditional rules respect priority, stop-if-true and blank/text values', () => {
  const first: ConditionalRule = { id: 0, range: 'A1:B2', operator: 'greaterThan', value: 5, secondValue: null,
    color: '#ff0000', background: null, priority: 1, stopIfTrue: false, supported: true };
  const second = { ...first, id: 1, color: '#000000', background: '#ffff00', priority: 2 };
  const cell = { address: 'A1', type: 'n', value: '12', formula: null, style: 0 };
  expect(conditionalStyle(cell, defaultStyle, [first, second])).toMatchObject({ color: '#ff0000', background: '#ffff00' });
  expect(conditionalStyle(cell, defaultStyle, [{ ...first, stopIfTrue: true }, second]).background).toBe(defaultStyle.background);
  expect(conditionalStyle({ ...cell, type: 'inlineStr' }, defaultStyle, [second])).toBe(defaultStyle);
  expect(conditionalStyle({ ...cell, value: '' }, defaultStyle, [second])).toBe(defaultStyle);
  expect(conditionalStyle({ ...cell, address: 'C1' }, defaultStyle, [second]).background).toBe(defaultStyle.background);
});

test('create, render, update values, remove and undo a conditional rule', async ({ page, request }) => {
  const created = await request.post('/api/documents/new');
  const prepared = await request.post('/api/documents/edit', { multipart: {
    file: { name: 'conditional.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await created.body() },
    edit: JSON.stringify({ action: 'cells', sheetId: '1', cells: [{ address: 'A1', value: '12', type: 'number' }, { address: 'A2', value: '5', type: 'number' }] })
  } });
  await page.goto('/');
  await page.getByLabel('Open XLSX', { exact: true }).setInputFiles({ name: 'conditional.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await prepared.body() });
  const cell = (name: string) => page.getByRole('gridcell', { name, exact: true });
  await page.getByRole('button', { name: 'Conditional formatting', exact: true }).click();
  await page.getByLabel('Applies to', { exact: true }).fill('A1:A2');
  await page.getByLabel('Rule value', { exact: true }).fill('10');
  await page.getByRole('button', { name: 'Add rule', exact: true }).click();
  await expect(page.getByRole('heading', { name: 'Worksheet rules (1)' })).toBeVisible();
  await page.getByRole('button', { name: 'Done', exact: true }).click();
  await expect(cell('A1')).toHaveCSS('background-color', 'rgb(255, 199, 206)');
  await expect(cell('A2')).not.toHaveCSS('background-color', 'rgb(255, 199, 206)');
  await cell('A2').dblclick();
  await page.getByRole('textbox', { name: 'Edit A2', exact: true }).fill('20');
  await page.keyboard.press('Enter');
  await expect(cell('A2')).toHaveCSS('background-color', 'rgb(255, 199, 206)');
  await page.getByRole('button', { name: 'Conditional formatting', exact: true }).click();
  await page.getByRole('button', { name: 'Remove rule 1', exact: true }).click();
  await expect(page.getByRole('heading', { name: 'Worksheet rules (0)' })).toBeVisible();
  await page.getByRole('button', { name: 'Done', exact: true }).click();
  await expect(cell('A1')).not.toHaveCSS('background-color', 'rgb(255, 199, 206)');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(cell('A1')).toHaveCSS('background-color', 'rgb(255, 199, 206)');
});
