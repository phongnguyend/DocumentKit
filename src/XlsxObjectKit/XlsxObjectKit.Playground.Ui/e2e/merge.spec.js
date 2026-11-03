import { test, expect } from '@playwright/test';

test('merge and unmerge preserve values with undo and reject data loss', async ({ page, request }) => {
  const created = await request.post('/api/documents/new');
  const prepared = await request.post('/api/documents/edit', { multipart: {
    file: { name: 'merge.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await created.body() },
    edit: JSON.stringify({ action: 'cells', sheetId: '1', cells: [{ address: 'A1', value: 'Title' }, { address: 'D1', value: 'Keep me' }] })
  } });
  await page.goto('/');
  await page.getByLabel('Open XLSX', { exact: true }).setInputFiles({ name: 'merge.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await prepared.body() });
  const cell = name => page.getByRole('gridcell', { name, exact: true });
  async function select(from, to) {
    const a = await cell(from).boundingBox();
    const b = await cell(to).boundingBox();
    await page.mouse.move(a.x + 10, a.y + 10);
    await page.mouse.down();
    await page.mouse.move(b.x + 10, b.y + 10, { steps: 6 });
    await page.mouse.up();
  }
  await select('A1', 'C2');
  await page.getByRole('button', { name: 'Merge cells', exact: true }).click();
  await expect(cell('B1')).toHaveCount(0);
  await expect(cell('A1')).toHaveText('Title');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(cell('B1')).toBeVisible();
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(cell('B1')).toHaveCount(0);
  await page.getByRole('button', { name: 'Unmerge cells', exact: true }).click();
  await expect(cell('B1')).toBeVisible();
  await expect(cell('A1')).toHaveText('Title');
  await select('A1', 'D2');
  await page.getByRole('button', { name: 'Merge cells', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('Only the top-left cell');
  await expect(cell('D1')).toHaveText('Keep me');
  await expect(cell('B1')).toBeVisible();
});
