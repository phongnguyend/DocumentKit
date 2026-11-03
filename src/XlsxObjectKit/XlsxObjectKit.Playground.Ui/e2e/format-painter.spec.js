import { test, expect } from '@playwright/test';
import { execFileSync } from 'node:child_process';
import { resolve } from 'node:path';

test('format painter copies formatting only, supports undo, repeated painting and Escape', async ({ page }, info) => {
  const fixture = info.outputPath('painter.xlsx');
  execFileSync('dotnet', ['run', '--project', resolve('../XlsxObjectKit.Tests'), '-c', 'Release', '--no-build', '--', '--write-formatting-fixture', fixture]);
  await page.goto('/');
  await page.getByLabel('Open XLSX', { exact: true }).setInputFiles(fixture);
  const painter = page.getByRole('button', { name: 'Format Painter', exact: true });
  const a1 = page.getByRole('gridcell', { name: 'A1', exact: true });
  const a2 = page.getByRole('gridcell', { name: 'A2', exact: true });
  await a1.click();
  const grid = page.getByRole('grid');
  const bounds = await grid.boundingBox();
  await painter.click();
  await expect(painter).toHaveAttribute('aria-pressed', 'true');
  expect(await grid.boundingBox()).toEqual(bounds);
  await a2.click();
  await expect(a2).toHaveText('0.13');
  await expect(a2).toHaveCSS('font-weight', '700');
  await expect(a2).toHaveCSS('border-bottom-style', 'double');
  await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveValue('0.125');
  await expect(painter).toHaveAttribute('aria-pressed', 'false');
  expect(await grid.boundingBox()).toEqual(bounds);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(a2).toHaveText('12.50%');
  await expect(a2).toHaveCSS('font-weight', '400');
  await a1.click();
  await painter.dblclick();
  await expect(page.locator('.painter-status')).toContainText('Keep painting');
  await a2.click();
  await expect(a2).toHaveCSS('font-weight', '700');
  await expect(painter).toHaveAttribute('aria-pressed', 'true');
  const a3 = page.getByRole('gridcell', { name: 'A3', exact: true });
  await a3.click();
  await expect(a3).toHaveText('0.00');
  await expect(a3).toHaveCSS('font-weight', '700');
  await page.keyboard.press('Escape');
  await expect(painter).toHaveAttribute('aria-pressed', 'false');
  const b4 = page.getByRole('gridcell', { name: 'B4', exact: true });
  await b4.click();
  await expect(b4).toHaveCSS('font-weight', '400');
});


test('drag paints a rectangular range once and Escape cancels a pending drag', async ({ page }, info) => {
  const fixture = info.outputPath('range.xlsx');
  execFileSync('dotnet', ['run', '--project', resolve('../XlsxObjectKit.Tests'), '-c', 'Release', '--no-build', '--', '--write-formatting-fixture', fixture]);
  await page.goto('/');
  await page.getByLabel('Open XLSX', { exact: true }).setInputFiles(fixture);
  const cell = name => page.getByRole('gridcell', { name, exact: true });
  await expect(cell('A1')).toBeVisible();
  const original = {};
  for (const name of ['A2', 'B2', 'A3', 'B3']) {
    original[name] = await cell(name).evaluate(element => ({ weight: getComputedStyle(element).fontWeight, alignment: getComputedStyle(element).textAlign }));
  }
  const painter = page.getByRole('button', { name: 'Format Painter', exact: true });
  await cell('A1').click();
  await painter.click();
  const start = await cell('B3').boundingBox();
  const end = await cell('A2').boundingBox();
  await page.mouse.move(start.x + start.width / 2, start.y + start.height / 2);
  await page.mouse.down();
  await page.mouse.move(end.x + end.width / 2, end.y + end.height / 2, { steps: 8 });
  await expect(page.locator('.paint-range')).toBeVisible();
  await expect(cell('A2')).toHaveText('12.50%');
  await page.mouse.up();
  for (const name of ['A2', 'B2', 'A3', 'B3']) {
    await expect(cell(name)).toHaveCSS('font-weight', '700');
    await expect(cell(name)).toHaveCSS('text-align', 'right');
  }
  await expect(cell('A2')).toHaveText('0.13');
  await expect(painter).toHaveAttribute('aria-pressed', 'false');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(cell('A2')).toHaveText('12.50%');
  for (const name of ['A2', 'B2', 'A3', 'B3']) {
    await expect(cell(name)).toHaveCSS('font-weight', original[name].weight);
    await expect(cell(name)).toHaveCSS('text-align', original[name].alignment);
  }
  await cell('A1').click();
  await painter.click();
  await page.mouse.move(start.x + start.width / 2, start.y + start.height / 2);
  await page.mouse.down();
  await page.mouse.move(end.x + end.width / 2, end.y + end.height / 2, { steps: 8 });
  await page.keyboard.press('Escape');
  await page.mouse.up();
  await expect(page.locator('.paint-range')).toHaveCount(0);
  await expect(painter).toHaveAttribute('aria-pressed', 'false');
  await expect(cell('A2')).toHaveText('12.50%');
  await expect(page.getByRole('button', { name: 'Redo', exact: true })).toBeEnabled();
});
