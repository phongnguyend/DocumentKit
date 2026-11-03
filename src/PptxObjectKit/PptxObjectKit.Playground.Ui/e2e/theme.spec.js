import { test, expect } from '@playwright/test';

test('theme persists, follows system preferences, and syncs browser tabs', async ({ page, context }) => {
  await page.emulateMedia({ colorScheme: 'light' });
  await page.goto('/');
  const button = name => page.getByRole('button', { name: name + ' theme', exact: true });
  await expect(button('System')).toHaveAttribute('aria-pressed', 'true');
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'light');
  await button('Dark').click();
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
  await expect(page.locator('header').first()).toHaveCSS('background-color', 'rgb(27, 37, 53)');
  await expect(button('Dark')).toHaveCSS('background-color', 'rgb(49, 85, 129)');
  await page.reload();
  await expect(button('Dark')).toHaveAttribute('aria-pressed', 'true');
  const other = await context.newPage();
  await other.goto(page.url());
  await other.getByRole('button', { name: 'Light theme', exact: true }).click();
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'light');
  await expect(button('Light')).toHaveAttribute('aria-pressed', 'true');
  await other.close();
  await page.emulateMedia({ colorScheme: 'dark' });
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'light');
  await button('System').click();
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
  await page.screenshot({ path: test.info().outputPath('dark.png') });
  await page.emulateMedia({ colorScheme: 'light' });
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'light');
  await page.setViewportSize({ width: 390, height: 844 });
  await expect(button('Dark')).toBeInViewport();
});

test('theme works when browser storage is unavailable', async ({ page }) => {
  await page.addInitScript(() => {
    Storage.prototype.getItem = () => {
      throw new Error('Blocked');
    };
    Storage.prototype.setItem = () => {
      throw new Error('Blocked');
    };
  });
  await page.goto('/');
  await page.getByRole('button', { name: 'Dark theme', exact: true }).click();
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
  await page.getByRole('button', { name: 'Light theme', exact: true }).click();
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'light');
});
