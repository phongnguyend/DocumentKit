import { test, expect } from '@playwright/test';
import { readFile } from 'node:fs/promises';

test('inspector selects stored cell types and preserves text and booleans during editing', async ({ page, request }) => {
  const created = await request.post('/api/documents/new');
  const prepared = await request.post('/api/documents/edit', { multipart: {
    file: { name: 'types.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await created.body() },
    edit: JSON.stringify({ action: 'cells', sheetId: '1', cells: [
      { address: 'A1', value: '00123', type: 'text' },
      { address: 'A2', value: '0', type: 'number' },
      { address: 'A3', value: 'true', type: 'boolean' },
      { address: 'A4', value: '="text result"', type: 'formula' },
      { address: 'A5', value: '=1/0', type: 'formula' }
    ] })
  } });
  expect(prepared.ok()).toBeTruthy();
  await page.goto('/');
  await page.getByLabel('Open XLSX', { exact: true }).setInputFiles({ name: 'types.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await prepared.body() });
  const type = page.getByRole('combobox', { name: 'Value type', exact: true });
  for (const [address, expected] of [['A1', 'text'], ['A2', 'number'], ['A3', 'boolean'], ['A4', 'formula'], ['A5', 'formula'], ['A6', 'auto']]) {
    await page.getByRole('gridcell', { name: address, exact: true }).click();
    await expect(type).toHaveValue(expected);
  }
  await page.getByRole('gridcell', { name: 'A1', exact: true }).click();
  await page.getByRole('textbox', { name: 'Formula bar' }).fill('00456');
  await page.getByRole('button', { name: 'Apply cell value' }).click();
  await expect(page.getByRole('gridcell', { name: 'A1', exact: true })).toHaveText('00456');
  await page.getByRole('gridcell', { name: 'A1', exact: true }).dblclick();
  await page.getByRole('textbox', { name: 'Edit A1', exact: true }).fill('00789');
  await page.getByRole('textbox', { name: 'Edit A1', exact: true }).press('Enter');
  await expect(page.getByRole('gridcell', { name: 'A1', exact: true })).toHaveText('00789');
  await page.getByRole('gridcell', { name: 'A3', exact: true }).dblclick();
  await expect(page.getByRole('textbox', { name: 'Edit A3', exact: true })).toHaveValue('true');
  await page.getByRole('textbox', { name: 'Edit A3', exact: true }).fill('false');
  await page.getByRole('textbox', { name: 'Edit A3', exact: true }).press('Enter');
  await expect(page.getByRole('gridcell', { name: 'A3', exact: true })).toHaveText('FALSE');
  await expect(type).toHaveValue('boolean');
  await page.getByRole('gridcell', { name: 'A3', exact: true }).click();
  await page.getByRole('grid').press('Delete');
  await expect(type).toHaveValue('auto');
  await expect(page.getByRole('gridcell', { name: 'A3', exact: true })).toHaveText('');
});

test('always keeps a default upload tab', async ({ page }) => {
  await page.goto('/');
  await expect(page.getByRole('tab')).toHaveCount(1);
  await expect(page.getByRole('tab', { name: 'New tab', exact: true })).toHaveAttribute('aria-selected', 'true');
  await page.getByRole('button', { name: 'Close New tab', exact: true }).click();
  await expect(page.getByRole('tab')).toHaveCount(1);
  await expect(page.getByRole('tabpanel', { name: 'Upload workbook' })).toBeVisible();
  await page.getByRole('button', { name: 'New workbook', exact: true }).click();
  await expect(page.getByRole('grid')).toBeVisible();
  await expect(page.getByRole('tab')).toHaveCount(1);
  await page.getByRole('button', { name: 'Close Workbook.xlsx', exact: true }).click();
  await expect(page.getByRole('tab')).toHaveCount(1);
  await expect(page.getByRole('tab', { name: 'New tab', exact: true })).toHaveAttribute('aria-selected', 'true');
  await expect(page.getByRole('tabpanel', { name: 'Upload workbook' })).toBeVisible();
});

test('new tab opens an upload screen and is replaced by the uploaded workbook', async ({ page, request }) => {
  const response = await request.post('/api/documents/new');
  expect(response.ok()).toBeTruthy();
  const buffer = await response.body();
  await page.goto('/');
  await page.getByRole('button', { name: 'New workbook', exact: true }).click();
  await expect(page.getByRole('grid')).toBeVisible();
  await page.getByRole('textbox', { name: 'Formula bar' }).fill('Keep this edit');
  await page.getByRole('button', { name: 'Apply cell value' }).click();
  await expect(page.getByRole('gridcell', { name: 'A1', exact: true })).toHaveText('Keep this edit');
  await page.getByRole('button', { name: 'New tab', exact: true }).click();
  await expect(page.getByRole('tab', { name: 'New tab', exact: true })).toHaveAttribute('aria-selected', 'true');
  await expect(page.getByRole('grid')).toBeHidden();
  await page.getByLabel('Choose XLSX file', { exact: true }).setInputFiles({ name: 'Uploaded.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer });
  await expect(page.getByRole('tab', { name: 'Uploaded.xlsx', exact: true })).toHaveAttribute('aria-selected', 'true');
  await expect(page.getByRole('tab')).toHaveCount(2);
  await page.getByRole('tab', { name: 'Workbook.xlsx *', exact: true }).click();
  await expect(page.getByRole('gridcell', { name: 'A1', exact: true })).toHaveText('Keep this edit');
  await page.getByRole('button', { name: 'New tab', exact: true }).click();
  await page.getByRole('button', { name: 'Close New tab', exact: true }).click();
  await expect(page.getByRole('dialog')).toHaveCount(0);
  await expect(page.getByRole('tab')).toHaveCount(2);
});

test('create, edit, format, manage worksheets, download and preserve separate tabs', async ({ page, request }, info) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'New workbook', exact: true }).click();
  const grid = page.getByRole('grid');
  const a1 = grid.getByRole('gridcell', { name: 'A1', exact: true });
  await a1.dblclick();
  await page.getByRole('textbox', { name: 'Edit A1', exact: true }).fill('Hello workbook');
  await page.getByRole('textbox', { name: 'Edit A1', exact: true }).press('Enter');
  await expect(a1).toHaveText('Hello workbook');
  await page.getByRole('button', { name: 'Bold', exact: true }).click();
  await page.getByRole('button', { name: 'Apply formatting', exact: true }).click();
  await expect(a1).toHaveCSS('font-weight', '700');
  await page.getByRole('spinbutton', { name: 'Row height (pt)' }).fill('40');
  await page.getByRole('button', { name: 'Apply sizes' }).click();
  await expect.poll(async () => (await a1.boundingBox()).height).toBeGreaterThan(50);
  await page.locator('summary').filter({ hasText: /^Worksheet$/ }).click();
  await page.getByRole('button', { name: 'Rename worksheet', exact: true }).click();
  await page.getByRole('textbox', { name: 'Worksheet name' }).fill('Budget');
  await page.getByRole('button', { name: 'Save worksheet' }).click();
  await expect(page.getByRole('grid')).toHaveAttribute('aria-label', 'Budget cells');
  await page.locator('.sheets-panel').getByRole('button', { name: 'Add worksheet', exact: true }).click();
  await page.getByRole('textbox', { name: 'Worksheet name' }).fill('Notes');
  await page.getByRole('button', { name: 'Save worksheet' }).click();
  await expect(page.locator('.sheets-panel').getByRole('button', { name: 'Notes', exact: true })).toBeVisible();
  await page.getByRole('combobox', { name: 'Value type', exact: true }).selectOption('formula');
  await page.getByRole('textbox', { name: 'Formula bar' }).fill('=1+2');
  await page.getByRole('button', { name: 'Apply cell value' }).click();
  await expect(a1).toHaveText('3');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(a1).toHaveText('Hello workbook');
  const downloadEvent = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download XLSX' }).click();
  const download = await downloadEvent;
  const bytes = await readFile(await download.path());
  const inspected = await request.post('/api/documents/inspect', { data: bytes });
  expect(inspected.ok()).toBeTruthy();
  const content = await inspected.json();
  expect(content.sheets.map(sheet => sheet.name)).toEqual(['Budget', 'Notes']);
  expect(content.sheets[0].cells[0].value).toBe('Hello workbook');
  await page.getByLabel('Open XLSX', { exact: true }).setInputFiles({ name: 'Reopened.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: bytes });
  await expect(page.getByRole('tab')).toHaveCount(2);
  await expect(page.getByRole('tab', { name: 'Reopened.xlsx' })).toHaveAttribute('aria-selected', 'true');
  await expect(page.getByRole('gridcell', { name: 'A1', exact: true })).toHaveText('Hello workbook');
  await page.getByRole('tab', { name: 'Workbook.xlsx', exact: true }).click();
  await expect(page.getByRole('gridcell', { name: 'A1', exact: true })).toHaveText('Hello workbook');
  await expect(page.getByRole('gridcell', { name: 'A20', exact: true })).toBeVisible();
  await page.locator('.workbook-panel:not([hidden]) .inspector').evaluate(element => element.scrollTop = 0);
  expect(await page.evaluate(() => document.documentElement.scrollHeight <= window.innerHeight)).toBeTruthy();
  await page.screenshot({ path: info.outputPath('xlsx-playground.png') });
});

test('paste a cell block and reach another virtual window', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'New workbook', exact: true }).click();
  await page.getByRole('grid').evaluate(element => {
    const clipboardData = new DataTransfer();
    clipboardData.setData('text/plain', 'Item\tQuantity\nPencils\t12');
    element.dispatchEvent(new ClipboardEvent('paste', { clipboardData, bubbles: true }));
  });
  await expect(page.getByRole('gridcell', { name: 'B2', exact: true })).toHaveText('12');
  await page.getByRole('textbox', { name: 'Go to cell' }).fill('AA10001');
  await page.getByRole('textbox', { name: 'Go to cell' }).press('Enter');
  await page.getByRole('textbox', { name: 'Formula bar' }).fill('Distant cell');
  await page.getByRole('button', { name: 'Apply cell value' }).click();
  await expect(page.getByRole('gridcell', { name: 'AA10001', exact: true })).toHaveText('Distant cell');
  await page.getByRole('button', { name: 'Close Workbook.xlsx' }).click();
  await expect(page.getByRole('dialog', { name: 'Close workbook?' })).toBeVisible();
  await page.getByRole('button', { name: 'Keep editing' }).click();
  await expect(page.getByRole('gridcell', { name: 'AA10001', exact: true })).toHaveText('Distant cell');
});

test('invalid uploads show a recoverable error', async ({ page }) => {
  await page.goto('/');
  await page.getByLabel('Open XLSX', { exact: true }).setInputFiles({ name: 'invalid.xlsx', mimeType: 'application/octet-stream', buffer: Buffer.from('not a zip') });
  await expect(page.getByRole('alert')).toBeVisible();
  await page.getByRole('button', { name: 'New workbook', exact: true }).click();
  await expect(page.getByRole('grid')).toBeVisible();
});
