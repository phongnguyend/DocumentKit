import { test, expect } from '@playwright/test';

test('scroll beyond Z, edit AA and navigate to the final worksheet column', async ({ page, request }) => {
  const created = await request.post('/api/documents/new');
  const file = (buffer: Buffer) => ({ name: 'wide.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer });
  const prepared = await request.post('/api/documents/edit', { multipart: {
    file: file(await created.body()), edit: JSON.stringify({ action: 'cells', sheetId: '1', cells: [
      { address: 'Z1', value: 'Last single letter' }, { address: 'AA1', value: 'Beyond Z' },
      { address: 'AQ1', value: 'Farther right' },
      { address: 'Z2', value: 'Merged across boundary' }
    ] })
  } });
  const merged = await request.post('/api/documents/edit', { multipart: {
    file: file(await prepared.body()), edit: JSON.stringify({ action: 'merge-cells', sheetId: '1', range: 'Z2:AB2' })
  } });
  expect(merged.ok()).toBeTruthy();
  await page.goto('/');
  await page.getByLabel('Open XLSX', { exact: true }).setInputFiles(file(await merged.body()));
  const grid = page.getByRole('grid');
  expect(await grid.evaluate(element => element.scrollWidth)).toBeLessThan(4000);
  const cell = (name: string) => page.getByRole('gridcell', { name, exact: true });
  await grid.evaluate(element => { element.scrollLeft = 1550; });
  await expect(page.getByRole('button', { name: 'Select column AA', exact: true })).toBeVisible();
  await expect(cell('AA1')).toHaveText('Beyond Z');
  await expect(cell('Z2')).toHaveText('Merged across boundary');
  await expect(cell('AA2')).toHaveCount(0);
  await cell('Z1').click();
  await page.keyboard.press('ArrowRight');
  await expect(cell('AA1')).toHaveAttribute('aria-selected', 'true');
  await cell('AA1').dblclick();
  await page.getByRole('textbox', { name: 'Edit AA1', exact: true }).fill('Edited beyond Z');
  await page.keyboard.press('Enter');
  await expect(cell('AA1')).toHaveText('Edited beyond Z');
  await grid.evaluate(element => { element.scrollLeft = 2500; });
  await expect(cell('AQ1')).toBeVisible();
  const go = page.getByRole('textbox', { name: 'Go to cell', exact: true });
  await go.fill('XFD1');
  await go.press('Enter');
  await expect(page.getByRole('button', { name: 'Select column XFD', exact: true })).toBeVisible();
  await expect(cell('XFD1')).toHaveText('');
  expect(await page.getByRole('columnheader').count()).toBeLessThan(50);
  await go.fill('A1');
  await go.press('Enter');
  await expect(page.getByRole('button', { name: 'Select column A', exact: true })).toBeVisible();
  expect(await grid.evaluate(element => element.scrollWidth)).toBeLessThan(4000);
});
