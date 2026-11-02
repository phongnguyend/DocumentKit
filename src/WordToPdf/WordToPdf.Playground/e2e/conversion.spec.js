import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';

const file = { name: 'sample.docx', mimeType: 'application/octet-stream', buffer: readFileSync(new URL('./fixtures/sample.docx', import.meta.url)) };

test('validates uploads and accepts drag and drop', async ({ page }) => {
  await page.goto('/');
  await expect(page.getByRole('button', { name: 'Convert to PDF' })).toBeDisabled();
  for (const [upload, message] of [
    [{ ...file, name: 'wrong.html' }, 'extension'],
    [{ ...file, buffer: Buffer.alloc(0) }, 'empty'],
    [{ ...file, buffer: Buffer.alloc(10 * 1024 * 1024 + 1) }, '10 MB'],
  ]) {
    await page.getByLabel('Word file').setInputFiles(upload);
    await expect(page.getByRole('alert')).toContainText(message);
  }
  const transfer = await page.evaluateHandle(() => {
    const data = new DataTransfer();
    data.items.add(new File(['test'], 'sample.DOCX'));
    return data;
  });
  await page.locator('.dropzone').dispatchEvent('drop', { dataTransfer: transfer });
  await expect(page.getByRole('button', { name: 'Convert to PDF' })).toBeEnabled();
  await page.getByRole('checkbox', { name: /LibreOffice/ }).uncheck();
  await expect(page.getByRole('button', { name: 'Convert to PDF' })).toBeDisabled();
});

test('converts through the proxy route and provides view and download links', async ({ page }) => {
  await page.route('**/api/libreoffice/convert', async (route) => {
    expect(route.request().method()).toBe('POST');
    expect(route.request().postDataBuffer().toString()).toContain('name="file"; filename="sample.docx"');
    await route.fulfill({ contentType: 'application/pdf', body: '%PDF-1.4\n%%EOF' });
  });
  await page.setViewportSize({ width: 1280, height: 720 });
  await page.goto('/');
  await page.getByLabel('Word file').setInputFiles(file);
  await page.getByRole('button', { name: 'Convert to PDF' }).click();
  const view = page.getByRole('link', { name: 'View PDF' });
  await expect(view).toHaveAttribute('target', '_blank');
  const downloadPromise = page.waitForEvent('download');
  await page.getByRole('link', { name: 'Download PDF' }).click();
  expect((await downloadPromise).suggestedFilename()).toBe('sample.pdf');
  expect(await page.evaluate(() => document.documentElement.scrollHeight)).toBe(720);
  await page.screenshot({ path: 'test-results/desktop.png', fullPage: true });
  await page.setViewportSize({ width: 390, height: 844 });
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBe(390);
  await page.screenshot({ path: 'test-results/mobile.png', fullPage: true });
  await page.getByLabel('Word file').setInputFiles({ ...file, name: 'next.doc' });
  await expect(view).toHaveCount(0);
});

test('shows conversion errors and allows retry', async ({ page }) => {
  await page.route('**/api/libreoffice/convert', (route) => route.fulfill({ status: 422, contentType: 'application/problem+json', body: JSON.stringify({ detail: 'Cannot convert document.' }) }));
  await page.goto('/');
  await page.getByLabel('Word file').setInputFiles(file);
  await page.getByRole('button', { name: 'Convert to PDF' }).click();
  await expect(page.getByRole('alert')).toContainText('Cannot convert document.');
  await expect(page.getByRole('button', { name: 'Convert to PDF' })).toBeEnabled();
  await page.route('**/api/libreoffice/convert', (route) => route.fulfill({ contentType: 'text/html', body: '<html>Error</html>' }));
  await page.getByRole('button', { name: 'Convert to PDF' }).click();
  await expect(page.getByRole('alert')).toContainText('did not return a PDF');
});

test('real LibreOffice conversion and API validation through proxy', async ({ page, request }) => {
  test.setTimeout(120000);
  test.skip(process.env.WORD_TO_PDF_LIVE_TEST !== '1', 'Requires the proxy and LibreOffice API.');
  await page.goto('/');
  await page.getByLabel('Word file').setInputFiles(file);
  await page.getByRole('button', { name: 'Convert to PDF' }).click();
  const download = page.getByRole('link', { name: 'Download PDF' });
  await expect(download).toBeVisible({ timeout: 100000 });
  const signature = await download.evaluate(async (link) => {
    const response = await fetch(link.href);
    return (await response.blob()).slice(0, 5).text();
  });
  expect(signature).toBe('%PDF-');
  for (const [name, buffer, status] of [
    ['wrong.html', Buffer.from('wrong format'), 400],
    ['empty.docx', Buffer.alloc(0), 400],
    ['invalid.docx', Buffer.from('not a document'), 422],
  ]) {
    const response = await request.post('/api/libreoffice/convert', { multipart: { file: { name, mimeType: 'application/octet-stream', buffer } } });
    expect(response.status()).toBe(status);
  }
});


test('selects native conversion and shows conversion notes', async ({ page }) => {
  await page.route('**/api/docxobjectkit/convert', (route) => route.fulfill({ contentType: 'application/pdf', headers: { 'X-Conversion-Warning-Count': '2' }, body: '%PDF-1.4\n%%EOF' }));
  await page.goto('/');
  await page.getByRole('checkbox', { name: /LibreOffice/ }).uncheck();
  await page.getByRole('checkbox', { name: /DocxObjectKit/ }).check();
  await page.getByLabel('Word file').setInputFiles(file);
  await page.getByRole('button', { name: 'Convert to PDF' }).click();
  await expect(page.getByRole('region', { name: 'DocxObjectKit result' })).toContainText('2 conversion notes');
  await expect(page.getByRole('link', { name: 'Download PDF' })).toHaveAttribute('download', 'sample.pdf');
  await page.getByLabel('Word file').setInputFiles({ ...file, name: 'legacy.doc' });
  await page.getByRole('button', { name: 'Convert to PDF' }).click();
  await expect(page.getByRole('alert')).toContainText('Use LibreOffice for legacy .doc');
});

test('real native conversion and invalid uploads through proxy', async ({ request }) => {
  test.skip(process.env.WORD_TO_PDF_NATIVE_LIVE_TEST !== '1', 'Requires the native API and proxy.');
  const endpoint = (process.env.WORD_TO_PDF_NATIVE_URL || 'http://localhost:62823') + '/docxobjectkit/convert';
  const response = await request.post(endpoint, { multipart: { file } });
  expect(response.status()).toBe(200);
  expect(response.headers()['content-type']).toContain('application/pdf');
  expect(response.headers()['x-conversion-warning-count']).toMatch(/^\d+$/);
  expect((await response.body()).subarray(0, 5).toString()).toBe('%PDF-');
  for (const [name, buffer, status] of [
    ['legacy.doc', file.buffer, 400],
    ['empty.docx', Buffer.alloc(0), 400],
    ['invalid.docx', Buffer.from('not a zip archive'), 422],
  ]) {
    const invalid = await request.post(endpoint, { multipart: { file: { name, buffer, mimeType: file.mimeType } } });
    expect(invalid.status()).toBe(status);
  }
});
