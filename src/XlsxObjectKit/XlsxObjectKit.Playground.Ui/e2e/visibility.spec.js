import { test, expect } from '@playwright/test';

test('hide selected headers, unhide and undo preserve worksheet data', async ({ page, request }) => {
  const created = await request.post('/api/documents/new');
  const prepared = await request.post('/api/documents/edit', { multipart: {
    file: { name: 'visibility.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await created.body() },
    edit: JSON.stringify({ action: 'cells', sheetId: '1', cells: [{ address: 'B2', value: 'Keep me' }] })
  } });
  await page.goto('/');
  await page.getByLabel('Open XLSX', { exact: true }).setInputFiles({ name: 'visibility.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await prepared.body() });
  async function action(name) {
    await page.locator('.toolbar-menu summary').filter({ hasText: /^Visibility$/ }).click();
    if (name.endsWith('columns')) {
      await expect(page.getByRole('button', { name: /^(Hide selected|Unhide selected) rows$/ })).toHaveCount(0);
    } else {
      await expect(page.getByRole('button', { name: /^(Hide selected|Unhide selected) columns$/ })).toHaveCount(0);
    }
    await page.getByRole('button', { name, exact: true }).click();
  }
  await page.getByRole('button', { name: 'Select column B', exact: true }).click();
  await action('Hide selected columns');
  await expect(page.getByRole('button', { name: 'Select column B', exact: true })).toHaveCount(0);
  await expect(page.getByRole('gridcell', { name: 'B2', exact: true })).toHaveCount(0);
  await page.getByRole('grid').focus();
  await page.keyboard.press('Escape');
  await page.locator('.toolbar-menu summary').filter({ hasText: /^Visibility$/ }).click();
  await expect(page.getByRole('button', { name: /^(Hide selected|Unhide selected)/ })).toHaveCount(0);
  await page.getByRole('button', { name: 'Unhide all columns', exact: true }).click();
  await expect(page.getByRole('gridcell', { name: 'B2', exact: true })).toHaveText('Keep me');
  await page.getByRole('gridcell', { name: 'A1', exact: true }).click();
  await expect(page.locator('.toolbar-menu summary').filter({ hasText: /^Visibility$/ })).toBeVisible();
  await expect(page.locator('.toolbar-menu summary').filter({ hasText: /^Visibility$/ })).toHaveAttribute('aria-disabled', 'true');
  await page.getByRole('button', { name: 'Select row 2', exact: true }).click();
  await action('Hide selected rows');
  await expect(page.getByRole('button', { name: 'Select row 2', exact: true })).toHaveCount(0);
  await expect(page.getByRole('gridcell', { name: 'B2', exact: true })).toHaveCount(0);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.getByRole('gridcell', { name: 'B2', exact: true })).toHaveText('Keep me');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(page.getByRole('gridcell', { name: 'B2', exact: true })).toHaveCount(0);
  await page.getByRole('gridcell', { name: 'A1', exact: true }).click();
  await page.locator('.toolbar-menu summary').filter({ hasText: /^Visibility$/ }).click();
  await expect(page.getByRole('button', { name: /^(Hide selected|Unhide selected)/ })).toHaveCount(0);
  await page.getByRole('button', { name: 'Unhide all rows', exact: true }).click();
  await expect(page.getByRole('gridcell', { name: 'B2', exact: true })).toHaveText('Keep me');
});
