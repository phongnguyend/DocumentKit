import { test, expect } from '@playwright/test';

test('delete selected rows and columns shifts data, recalculates and supports undo', async ({ page, request }) => {
  const created = await request.post('/api/documents/new');
  const prepared = await request.post('/api/documents/edit', { multipart: {
    file: { name: 'delete.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await created.body() },
    edit: JSON.stringify({ action: 'cells', sheetId: '1', cells: [
      { address: 'A1', value: '1', type: 'number' }, { address: 'A2', value: '2', type: 'number' },
      { address: 'A3', value: '3', type: 'number' }, { address: 'A4', value: '4', type: 'number' },
      { address: 'B4', value: 'SUM(A1:A3)', type: 'formula' }, { address: 'D4', value: 'Keep' }
    ] })
  } });
  expect(prepared.ok()).toBeTruthy();
  await page.goto('/');
  await page.getByLabel('Open XLSX', { exact: true }).setInputFiles({ name: 'delete.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await prepared.body() });
  const cell = name => page.getByRole('gridcell', { name, exact: true });
  async function selectHeaders(axis, first, last) {
    const a = await page.getByRole('button', { name: `Select ${axis} ${first}`, exact: true }).boundingBox();
    const b = await page.getByRole('button', { name: `Select ${axis} ${last}`, exact: true }).boundingBox();
    await page.mouse.move(a.x + a.width / 2, a.y + a.height / 2);
    await page.mouse.down();
    await page.mouse.move(b.x + b.width / 2, b.y + b.height / 2, { steps: 4 });
    await page.mouse.up();
  }
  async function remove(axis) {
    await page.locator('.toolbar-menu summary').filter({ hasText: /^Delete$/ }).click();
    await page.getByRole('button', { name: `Delete selected ${axis}`, exact: true }).click();
  }
  await selectHeaders('row', 2, 3);
  await remove('rows');
  await expect(cell('A2')).toHaveText('4');
  await expect(cell('B2')).toHaveText('1');
  await expect(cell('D2')).toHaveText('Keep');
  await expect(cell('A4')).toHaveText('');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(cell('A4')).toHaveText('4');
  await expect(cell('B4')).toHaveText('6');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(cell('A2')).toHaveText('4');
  await selectHeaders('column', 'C', 'B');
  await remove('columns');
  await expect(cell('B2')).toHaveText('Keep');
  await expect(cell('D2')).toHaveText('');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(cell('D2')).toHaveText('Keep');
  await page.getByRole('button', { name: 'Select column A', exact: true }).click();
  await remove('columns');
  await expect(cell('A2')).toHaveText('#REF!');
});
