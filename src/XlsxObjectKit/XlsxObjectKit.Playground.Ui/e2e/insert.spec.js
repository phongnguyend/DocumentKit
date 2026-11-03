import { test, expect } from '@playwright/test';

test('insert rows and columns from selections with formulas, undo and redo', async ({ page, request }) => {
  const created = await request.post('/api/documents/new');
  const prepared = await request.post('/api/documents/edit', { multipart: {
    file: { name: 'insert.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await created.body() },
    edit: JSON.stringify({ action: 'cells', sheetId: '1', cells: [
      { address: 'A1', value: '2', type: 'number' }, { address: 'A2', value: '5', type: 'number' },
      { address: 'B2', value: 'SUM(A1:A2)', type: 'formula' }
    ] })
  } });
  expect(prepared.ok()).toBeTruthy();
  await page.goto('/');
  await page.getByLabel('Open XLSX', { exact: true }).setInputFiles({ name: 'insert.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await prepared.body() });
  const cell = name => page.getByRole('gridcell', { name, exact: true });
  async function insert(action) {
    await page.locator('.toolbar-menu summary').filter({ hasText: /^Insert$/ }).click();
    await page.getByRole('button', { name: action, exact: true }).click();
  }
  await page.getByRole('button', { name: 'Select row 2', exact: true }).click();
  await insert('Insert rows above');
  await expect(cell('A3')).toHaveText('5');
  await expect(cell('B3')).toHaveText('7');
  await expect(cell('A2')).toHaveText('');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(cell('A2')).toHaveText('5');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(cell('A3')).toHaveText('5');
  const a = await page.getByRole('button', { name: 'Select column A', exact: true }).boundingBox();
  const b = await page.getByRole('button', { name: 'Select column B', exact: true }).boundingBox();
  await page.mouse.move(a.x + a.width / 2, a.y + a.height / 2);
  await page.mouse.down();
  await page.mouse.move(b.x + b.width / 2, b.y + b.height / 2, { steps: 4 });
  await page.mouse.up();
  await insert('Insert columns left');
  await expect(cell('C3')).toHaveText('5');
  await expect(cell('D3')).toHaveText('7');
  await expect(cell('A3')).toHaveText('');
  await cell('C1').click();
  await insert('Insert columns right');
  await expect(cell('E3')).toHaveText('7');
  await cell('C1').click();
  await insert('Insert rows below');
  await expect(cell('C4')).toHaveText('5');
  await expect(cell('E4')).toHaveText('7');
});
