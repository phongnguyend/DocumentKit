import { test, expect } from '@playwright/test';

test('drag header borders to resize, undo and cancel with Escape', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'New workbook', exact: true }).click();
  const cell = page.getByRole('gridcell', { name: 'A1', exact: true });
  const original = await cell.boundingBox();
  async function drag(label, dx, dy, cancel = false) {
    const handle = await page.getByRole('separator', { name: label, exact: true }).boundingBox();
    const x = handle.x + handle.width / 2;
    const y = handle.y + handle.height / 2;
    await page.mouse.move(x, y);
    await page.mouse.down();
    await page.mouse.move(x + dx, y + dy, { steps: 6 });
    await expect(page.locator('.dimension-guide')).toBeVisible();
    if (cancel) {
      await page.keyboard.press('Escape');
    }
    await page.mouse.up();
    await expect(page.locator('.dimension-guide')).toHaveCount(0);
  }
  await drag('Resize column A', 70, 0);
  await expect.poll(async () => Math.round((await cell.boundingBox()).width - original.width)).toBe(70);
  await expect(page.getByRole('separator', { name: 'Resize row 1', exact: true })).toHaveAttribute('aria-valuenow', '15');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect.poll(async () => (await cell.boundingBox()).width).toBe(original.width);
  await drag('Resize row 1', 0, 40);
  await expect.poll(async () => Math.round((await cell.boundingBox()).height - original.height)).toBe(40);
  await expect.poll(async () => (await cell.boundingBox()).width).toBe(original.width);
  await drag('Resize column A', 80, 0, true);
  await expect.poll(async () => (await cell.boundingBox()).width).toBe(original.width);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect.poll(async () => (await cell.boundingBox()).height).toBe(original.height);
});
