import { test, expect } from '@playwright/test';
import { readFile } from 'node:fs/promises';

test('worksheet colors save, undo, clear and survive download', async ({ page, request }, info) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'New workbook', exact: true }).click();
  const sheet = page.locator('.sheets-panel').getByRole('button', { name: 'Sheet1', exact: true });
  async function open() {
    await page.locator('.toolbar-menu summary').filter({ hasText: /^Worksheet$/ }).click();
    await page.getByRole('button', { name: 'Worksheet color', exact: true }).click();
  }
  await open();
  await page.getByRole('button', { name: 'Color #16a34a', exact: true }).click();
  await page.screenshot({ path: info.outputPath('sheet-color.png') });
  await page.getByRole('button', { name: 'Save color', exact: true }).click();
  await expect(page.getByRole('dialog')).toHaveCount(0);
  await expect(sheet).toHaveCSS('box-shadow', 'rgb(22, 163, 74) 4px 0px 0px 0px inset');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(sheet).toHaveCSS('box-shadow', 'none');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(sheet).toHaveCSS('box-shadow', 'rgb(22, 163, 74) 4px 0px 0px 0px inset');
  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download XLSX', exact: true }).click();
  const file = await download;
  const inspected = await request.post('/api/documents/inspect', { data: await readFile((await file.path())!), headers: { 'Content-Type': 'application/octet-stream' } });
  expect((await inspected.json()).sheets[0].tabColor).toBe('#16a34a');
  await open();
  await expect(page.getByRole('button', { name: 'Color #16a34a', exact: true })).toHaveAttribute('aria-pressed', 'true');
  await page.getByRole('button', { name: 'No color', exact: true }).click();
  await page.getByRole('button', { name: 'Save color', exact: true }).click();
  await expect(sheet).toHaveCSS('box-shadow', 'none');
});
