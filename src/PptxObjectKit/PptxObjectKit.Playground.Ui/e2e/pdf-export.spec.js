import { test, expect } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

const fixture = new URL('../../PptxObjectKit.Tests/Fixtures/sample.pptx', import.meta.url);

test('native PDF export uses current edits and offers options, notes, preview and download', async ({ page }, testInfo) => {
  await page.goto('/');
  await page.getByLabel('Open PPTX', { exact: true }).setInputFiles({ name: 'sample.pptx', mimeType: 'application/octet-stream', buffer: await readFile(fixture) });
  await expect(page.getByRole('img', { name: 'Slide 1', exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Select objects', exact: true }).click();
  await page.locator('.text-layer span[data-object-id]').first().click();
  await page.getByLabel('Text', { exact: true }).fill('Native PDF edited title');
  await page.getByRole('button', { name: 'Save text' }).click();
  await expect(page.locator('.text-layer')).toContainText('Native PDF edited title');
  await page.locator('summary').filter({ hasText: /^Document$/ }).click();
  await page.getByRole('button', { name: 'Export PDF', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: 'Export PDF', exact: true });
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Convert to PDF', exact: true }).click();
  await expect(dialog.getByRole('region', { name: 'Conversion result' })).toContainText('2 PDF pages ready');
  await page.screenshot({ path: testInfo.outputPath('pdf-export.png') });
  const downloaded = page.waitForEvent('download');
  await dialog.getByRole('link', { name: 'Download PDF' }).click();
  const file = await downloaded;
  expect(file.suggestedFilename()).toBe('sample.pdf');
  const output = testInfo.outputPath('native.pdf');
  await file.saveAs(output);
  expect((await readFile(output)).subarray(0, 5).toString()).toBe('%PDF-');
  await dialog.getByRole('button', { name: 'Preview', exact: true }).click();
  await expect(dialog.locator('iframe[title="Converted PDF preview"]')).toBeVisible();
  await dialog.getByRole('button', { name: 'Hide preview', exact: true }).click();
  await page.getByRole('combobox', { name: 'Slides', exact: true }).selectOption('current');
  await expect(dialog.getByRole('link', { name: 'Download PDF' })).toHaveCount(0);
  await dialog.getByRole('button', { name: 'Convert to PDF', exact: true }).click();
  await expect(dialog.getByRole('region', { name: 'Conversion result' })).toContainText('1 PDF page ready');
  for (const size of [{ width: 1280, height: 600 }, { width: 390, height: 700 }]) {
    await page.setViewportSize(size);
    await expect(dialog.getByRole('link', { name: 'Download PDF' })).toBeInViewport();
    expect(await dialog.evaluate(element => element.scrollHeight <= element.clientHeight + 1)).toBe(true);
  }
  await page.keyboard.press('Escape');
  await expect(dialog).toHaveCount(0);
});

test('PDF API rejects invalid selections and reports unavailable fonts in strict mode', async ({ request }) => {
  const bytes = await readFile(fixture);
  for (const slides of ['0', '3', '1,1', 'bad']) {
    const response = await request.post(`/api/documents/pdf?slides=${slides}`, { data: bytes });
    expect(response.status()).toBe(422);
  }
  const zip = await JSZip.loadAsync(bytes);
  const path = 'ppt/slides/slide1.xml';
  let xml = await zip.file(path).async('string');
  xml = xml.replace(/<a:latin[^>]*\/>/g, '').replace(/<a:rPr([^>]*?)\/>/g, '<a:rPr$1><a:latin typeface="UnavailableNativePdfFont"/></a:rPr>');
  zip.file(path, xml);
  const modified = await zip.generateAsync({ type: 'nodebuffer' });
  const response = await request.post('/api/documents/pdf?strict=true', { data: modified });
  expect(response.status()).toBe(422);
  expect((await response.json()).detail).toContain('Strict PDF conversion rejected');
});
