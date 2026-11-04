import { test, expect } from '@playwright/test';

test('worksheet bar sits below the expanded grid and scrolls on narrow screens', async ({ page, request }, info) => {
  let bytes = await (await request.post('/api/documents/new')).body();
  for (let i = 2; i <= 8; i++) {
    const response = await request.post('/api/documents/edit', { multipart: {
      file: { name: 'tabs.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: bytes },
      edit: JSON.stringify({ action: 'add-sheet', sheetId: '1', name: `Planning worksheet ${i}` })
    } });
    expect(response.ok()).toBeTruthy();
    bytes = await response.body();
    if (i <= 4) {
      const colored = await request.post('/api/documents/edit', { multipart: {
        file: { name: 'tabs.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: bytes },
        edit: JSON.stringify({ action: 'set-sheet-color', sheetId: String(i), tabColor: ['#ffdf57', '#e60000', '#0078d4'][i - 2] })
      } });
      expect(colored.ok()).toBeTruthy();
      bytes = await colored.body();
    }
  }
  await page.goto('/');
  await page.getByLabel('Open XLSX', { exact: true }).setInputFiles({ name: 'tabs.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: bytes });
  const bar = page.getByRole('navigation', { name: 'Worksheets', exact: true });
  const grid = page.getByRole('grid');
  await expect(bar).toBeVisible();
  const gridBox = (await grid.boundingBox())!;
  const barBox = (await bar.boundingBox())!;
  expect(barBox.y).toBeGreaterThanOrEqual(gridBox.y + gridBox.height - 1);
  expect(barBox.x).toBe(gridBox.x);
  expect(gridBox.x).toBeLessThan(10);
  await expect(bar.getByRole('button', { name: 'Add worksheet', exact: true })).toBeVisible();
  await expect(bar.getByRole('button', { name: 'Planning worksheet 3', exact: true })).toHaveCSS('background-color', 'rgb(230, 0, 0)');
  await expect(bar.getByRole('button', { name: 'Planning worksheet 3', exact: true })).toHaveCSS('color', 'rgb(255, 255, 255)');
  await page.screenshot({ path: info.outputPath('worksheet-bar.png') });
  await page.setViewportSize({ width: 720, height: 800 });
  await expect(bar).toBeVisible();
  const tabs = bar.locator('.sheet-tabs');
  expect(await tabs.evaluate(element => element.scrollWidth > element.clientWidth)).toBe(true);
  const last = bar.getByRole('button', { name: 'Planning worksheet 8', exact: true });
  await last.click();
  await expect(last).toHaveAttribute('aria-current', 'page');
  await expect(page.getByRole('grid')).toHaveAttribute('aria-label', 'Planning worksheet 8 cells');
  await expect(bar.getByRole('button', { name: 'Add worksheet', exact: true })).toBeVisible();
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
});
