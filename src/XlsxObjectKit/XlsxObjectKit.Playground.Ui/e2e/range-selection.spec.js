import { test, expect } from '@playwright/test';

test('merged cells have one full-size selection outline', async ({ page }) => {
  await page.route('**/api/documents/inspect', route => route.fulfill({ json: {
    sheets: [{ id: '1', name: 'Sheet1', part: 'xl/worksheets/sheet1.xml', hidden: false,
      cells: [{ address: 'A1', value: 'Merged title', type: 'inlineStr', style: 0, formula: null }],
      merges: ['A1:C2'], rowHeights: {}, columnWidths: {} }], styles: [], warnings: []
  } }));
  await page.goto('/');
  await page.getByRole('button', { name: 'New workbook', exact: true }).click();
  const merged = page.getByRole('gridcell', { name: 'A1', exact: true });
  await merged.click();
  await expect(page.locator('.range-selection-status')).toContainText('A1:C2');
  expect(await page.locator('.selected-range').boundingBox()).toEqual(await merged.boundingBox());
  await expect(page.locator('.grid-cell.selected')).toHaveCount(0);
  await expect(page.getByRole('gridcell', { selected: true })).toHaveCount(1);
  await expect(page.getByRole('gridcell', { name: 'B1', exact: true })).toHaveCount(0);
  const from = await page.getByRole('gridcell', { name: 'D4', exact: true }).boundingBox();
  const to = await merged.boundingBox();
  await page.mouse.move(from.x + 10, from.y + 10);
  await page.mouse.down();
  await page.mouse.move(to.x + to.width - 10, to.y + 10, { steps: 5 });
  await page.mouse.up();
  await expect(page.locator('.range-selection-status')).toContainText('A1:D4');
  await expect(page.locator('.selected-range')).toHaveCount(1);
  await expect(page.locator('.grid-cell.selected')).toHaveCount(0);
});

test('drag selects a rectangle, copies values, clears once and preserves inline editing', async ({ page, request }) => {
  const created = await request.post('/api/documents/new');
  const prepared = await request.post('/api/documents/edit', { multipart: {
    file: { name: 'range.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await created.body() },
    edit: JSON.stringify({ action: 'cells', sheetId: '1', cells: [
      { address: 'A1', value: 'First' }, { address: 'B1', value: 'Second' },
      { address: 'A2', value: 'Third' }, { address: 'B2', value: 'Fourth' }
    ] })
  } });
  await page.goto('/');
  await page.getByLabel('Open XLSX', { exact: true }).setInputFiles({ name: 'range.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await prepared.body() });
  const cell = name => page.getByRole('gridcell', { name, exact: true });
  const from = await cell('B2').boundingBox();
  const to = await cell('A1').boundingBox();
  await page.mouse.move(from.x + 15, from.y + 10);
  await page.mouse.down();
  await page.mouse.move(to.x + 15, to.y + 10, { steps: 5 });
  await page.mouse.up();
  await expect(page.locator('.range-selection-status')).toContainText('A1:B2');
  for (const name of ['A1', 'B1', 'A2', 'B2']) {
    await expect(cell(name)).toHaveAttribute('aria-selected', 'true');
  }
  await expect(cell('C2')).toHaveAttribute('aria-selected', 'false');
  const copied = await page.getByRole('grid').evaluate(element => {
    const data = new DataTransfer();
    element.dispatchEvent(new ClipboardEvent('copy', { bubbles: true, clipboardData: data }));
    return data.getData('text/plain');
  });
  expect(copied).toBe('First\tSecond\nThird\tFourth');
  await page.keyboard.press('Delete');
  await expect(cell('A1')).toHaveText('');
  await expect(cell('B2')).toHaveText('');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(cell('A1')).toHaveText('First');
  await expect(cell('B2')).toHaveText('Fourth');
  await cell('C3').click();
  await expect(cell('A1')).toHaveAttribute('aria-selected', 'false');
  await cell('C3').dblclick();
  await expect(page.getByRole('textbox', { name: 'Edit C3', exact: true })).toBeVisible();
  await page.keyboard.press('Escape');
});
