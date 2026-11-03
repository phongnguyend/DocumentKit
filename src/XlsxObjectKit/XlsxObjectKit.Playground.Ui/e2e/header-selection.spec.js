import { test, expect } from '@playwright/test';

test('merged cells do not expand header selections into neighboring columns or rows', async ({ page }) => {
  await page.route('**/api/documents/inspect', route => route.fulfill({ json: {
    sheets: [{ id: '1', name: 'Sheet1', part: 'xl/worksheets/sheet1.xml', hidden: false,
      cells: [{ address: 'A10', value: 'Merged footer', type: 'inlineStr', style: 0, formula: null }],
      merges: ['A10:D10', 'E1:E4'], rowHeights: {}, columnWidths: {} }], styles: [], warnings: []
  } }));
  await page.goto('/');
  await page.getByRole('button', { name: 'New workbook', exact: true }).click();
  for (const column of ['A', 'B', 'C', 'D']) {
    const button = page.getByRole('button', { name: 'Select column ' + column, exact: true });
    await button.click();
    await expect(page.locator('.range-selection-status')).toContainText(`${column}1:${column}1048576`);
    await expect(page.getByRole('columnheader', { selected: true })).toHaveCount(1);
    const outline = await page.locator('.selected-range').boundingBox();
    const header = await page.getByRole('columnheader', { selected: true }).boundingBox();
    expect(outline.x).toBeCloseTo(header.x, 0);
    expect(outline.width).toBeCloseTo(header.width, 0);
  }
  for (const row of [1, 2, 3, 4]) {
    await page.getByRole('button', { name: 'Select row ' + row, exact: true }).click();
    await expect(page.locator('.range-selection-status')).toContainText(`A${row}:XFD${row}`);
    await expect(page.getByRole('rowheader', { selected: true })).toHaveCount(1);
  }
  async function dragHeaders(axis, from, to, expectedRange, count, cancel = false) {
    const first = await page.getByRole('button', { name: `Select ${axis} ${from}`, exact: true }).boundingBox();
    const last = await page.getByRole('button', { name: `Select ${axis} ${to}`, exact: true }).boundingBox();
    await page.mouse.move(first.x + first.width / 2, first.y + first.height / 2);
    await page.mouse.down();
    await page.mouse.move(last.x + last.width / 2, last.y + last.height / 2, { steps: 6 });
    await expect(page.getByRole(axis === 'column' ? 'columnheader' : 'rowheader', { selected: true })).toHaveCount(count);
    if (cancel) {
      await page.keyboard.press('Escape');
    }
    await page.mouse.up();
    if (cancel) {
      await expect(page.locator('.selected-range')).toHaveCount(0);
    } else {
      await expect(page.locator('.range-selection-status')).toContainText(expectedRange);
      await expect(page.getByRole(axis === 'column' ? 'columnheader' : 'rowheader', { selected: true })).toHaveCount(count);
    }
  }
  await dragHeaders('column', 'A', 'C', 'A1:C1048576', 3);
  await dragHeaders('column', 'D', 'B', 'B1:D1048576', 3);
  await dragHeaders('row', 1, 3, 'A1:XFD3', 3);
  await dragHeaders('row', 4, 2, 'A2:XFD4', 3);
  await dragHeaders('column', 'A', 'D', '', 4, true);
});

test('header clicks select entire rows and columns, clear sparse cells and cancel', async ({ page, request }) => {
  const created = await request.post('/api/documents/new');
  const prepared = await request.post('/api/documents/edit', { multipart: {
    file: { name: 'headers.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await created.body() },
    edit: JSON.stringify({ action: 'cells', sheetId: '1', cells: [
      { address: 'B2', value: 'Keep' }, { address: 'C2', value: 'Clear' },
      { address: 'C20000', value: 'Outside window' }
    ] })
  } });
  expect(prepared.ok()).toBeTruthy();
  await page.goto('/');
  await page.getByLabel('Open XLSX', { exact: true }).setInputFiles({ name: 'headers.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await prepared.body() });
  const cell = name => page.getByRole('gridcell', { name, exact: true });
  await cell('B2').click();
  await page.getByRole('button', { name: 'Select column C', exact: true }).click();
  await expect(page.locator('.range-selection-status')).toContainText('C1:C1048576');
  await expect(cell('C2')).toHaveAttribute('aria-selected', 'true');
  await expect(cell('B2')).toHaveAttribute('aria-selected', 'false');
  const inspected = page.waitForResponse(response => response.url().endsWith('/api/documents/inspect') && response.ok());
  await page.keyboard.press('Delete');
  const afterClear = await (await inspected).json();
  expect(afterClear.sheets[0].cells.find(cell => cell.address === 'C20000')?.value ?? '').toBe('');
  await expect(cell('C2')).toHaveText('');
  await expect(cell('B2')).toHaveText('Keep');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(cell('C2')).toHaveText('Clear');
  await page.getByRole('button', { name: 'Select row 2', exact: true }).click();
  await expect(page.locator('.range-selection-status')).toContainText('A2:XFD2');
  await expect(cell('A2')).toHaveAttribute('aria-selected', 'true');
  await expect(cell('D2')).toHaveAttribute('aria-selected', 'true');
  await expect(cell('C3')).toHaveAttribute('aria-selected', 'false');
  await page.keyboard.press('Escape');
  await expect(page.locator('.selected-range')).toHaveCount(0);
  const header = page.getByRole('button', { name: 'Select column B', exact: true });
  await header.focus();
  await page.keyboard.press('Enter');
  await expect(page.locator('.range-selection-status')).toContainText('B1:B1048576');
  await expect(page.getByRole('textbox', { name: /^Edit / })).toHaveCount(0);
  await cell('D3').click();
  await expect(cell('B2')).toHaveAttribute('aria-selected', 'false');
  await page.getByRole('grid').evaluate(element => { element.scrollTop = 2000; });
  await page.getByRole('button', { name: 'Select column A', exact: true }).click();
  await expect(page.locator('.range-selection-status')).toContainText('A1:A1048576');
  expect(await page.getByRole('grid').evaluate(element => element.scrollTop)).toBe(2000);
});
