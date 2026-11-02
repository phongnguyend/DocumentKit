import { test, expect } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
const fixture = new URL('../../DocxObjectKit.Tests/Fixtures/sample.docx', import.meta.url);
async function open(page, files = ['sample.docx']) {
  await page.goto('/');
  const buffer = await readFile(fixture);
  await page.getByLabel('Open DOCX', { exact: true }).setInputFiles(files.map(name => ({ name, mimeType: 'application/octet-stream', buffer })));
  await expect(page.getByRole('img', { name: 'Page 1', exact: true })).toBeVisible();
}
async function menu(page, action) {
  await page.locator('summary').filter({ hasText: /^Document$/ }).click();
  await page.getByRole('button', { name: action, exact: true }).click();
}

test('tabs preserve text edits, selection modes, history and page navigation', async ({ page }, info) => {
  await open(page, ['one.docx', 'two.docx']);
  await expect(page.getByRole('tab')).toHaveCount(2);
  await page.getByRole('button', { name: 'Select objects', exact: true }).click();
  await page.locator('.canvas-host:visible [data-object-id="p0"]').click();
  await page.getByLabel('Text', { exact: true }).fill('Edited document title');
  await page.getByRole('button', { name: 'Save text', exact: true }).click();
  await expect(page.locator('.canvas-host:visible .text-layer')).toContainText('Edited document title');
  await page.getByRole('tab', { name: /two.docx/ }).click();
  await expect(page.locator('.canvas-host:visible .text-layer')).toContainText('Document toolkit sample');
  await page.getByRole('tab', { name: /one.docx/ }).click();
  await expect(page.locator('.canvas-host:visible .text-layer')).toContainText('Edited document title');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.locator('.canvas-host:visible .text-layer')).toContainText('Document toolkit sample');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await page.getByRole('button', { name: 'Next page', exact: true }).click();
  await expect(page.getByRole('img', { name: 'Page 2', exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Previous page', exact: true }).click();
  await page.getByRole('button', { name: 'Select text', exact: true }).click();
  const selection = await page.locator('.canvas-host:visible .text-layer').evaluate(layer => {
    const range = document.createRange();
    range.selectNodeContents(layer);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
    return selection.toString();
  });
  expect(selection).toContain('Edited document title');
  expect(await page.evaluate(() => document.documentElement.scrollHeight <= innerHeight + 1)).toBe(true);
  await page.getByRole('combobox', { name: 'Zoom', exact: true }).selectOption('100');
  await expect(page.locator('.canvas-host:visible')).toHaveClass(/zoomed/);
  expect(await page.evaluate(() => document.documentElement.scrollHeight <= innerHeight + 1)).toBe(true);
  await page.screenshot({ path: info.outputPath('workspace.png') });
});

test('image drag and resize autosave and persist on DOCX download', async ({ page }) => {
  await open(page);
  await page.getByRole('button', { name: 'Select objects', exact: true }).click();
  await page.locator('.canvas-host:visible [data-object-id="i0"]').click();
  const mover = page.getByRole('group', { name: 'Move selected image', exact: true });
  const box = await mover.boundingBox();
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width / 2 + 25, box.y + box.height / 2 + 20, { steps: 5 });
  const saved = page.waitForResponse(response => response.url().endsWith('/edit'));
  await page.mouse.up();
  expect((await saved).ok()).toBe(true);
  await expect(page.getByRole('button', { name: 'Download DOCX', exact: true })).toBeEnabled();
  await page.locator('.canvas-host:visible [data-object-id="i0"]').click();
  const handle = await page.getByRole('button', { name: 'Resize image se', exact: true }).boundingBox();
  await page.mouse.move(handle.x + handle.width / 2, handle.y + handle.height / 2);
  await page.mouse.down();
  await page.mouse.move(handle.x + handle.width / 2 + 20, handle.y + handle.height / 2 + 10, { steps: 5 });
  const resized = page.waitForResponse(response => response.url().endsWith('/edit'));
  await page.mouse.up();
  expect((await resized).ok()).toBe(true);
  await expect(page.getByRole('button', { name: 'Download DOCX', exact: true })).toBeEnabled();

  const promise = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download DOCX', exact: true }).click();
  const file = await promise;
  const zip = await JSZip.loadAsync(await readFile(await file.path()));
  const xml = await zip.file('word/document.xml').async('string');
  expect(xml).toContain('anchor');
  expect(xml).toContain('positionH');
});

test('markdown options and native PDF preview use current document', async ({ page }, info) => {
  await open(page);
  await menu(page, 'View Markdown');
  const markdown = page.getByRole('dialog', { name: 'View Markdown' });
  await expect(markdown.locator('article')).toContainText('Document toolkit sample');
  await markdown.getByRole('button', { name: 'Source', exact: true }).click();
  await expect(markdown.locator('pre')).toContainText('| Feature | Status |');
  await markdown.getByRole('checkbox', { name: 'Images', exact: true }).uncheck();
  await markdown.getByRole('button', { name: 'Reapply options', exact: true }).click();
  await expect(markdown.locator('pre')).not.toContainText('data:image');
  await page.getByRole('button', { name: 'Close Markdown', exact: true }).click();
  await menu(page, 'Export PDF');
  const dialog = page.getByRole('dialog', { name: 'Export PDF', exact: true });
  await dialog.getByRole('button', { name: 'Convert to PDF', exact: true }).click();
  await expect(dialog.getByRole('region', { name: 'Conversion result' })).toContainText('2 PDF pages ready');
  await dialog.getByRole('button', { name: 'Preview', exact: true }).click();
  await expect(dialog.locator('iframe')).toBeVisible();
  await page.screenshot({ path: info.outputPath('pdf-preview.png') });
  const promise = page.waitForEvent('download');
  await dialog.getByRole('link', { name: 'Download PDF', exact: true }).click();
  expect((await readFile(await (await promise).path())).subarray(0, 5).toString()).toBe('%PDF-');
});

test('API rejects invalid documents and bounds and preserves shared pictures on replacement', async ({ request }) => {
  expect((await request.post('/api/documents/inspect', { data: Buffer.from('invalid') })).status()).toBe(422);
  const buffer = await readFile(fixture);
  const source = await JSZip.loadAsync(buffer);
  const imageBase64 = await source.file('word/media/sample.png').async('base64');
  async function edit(value) {
    return request.post('/api/documents/edit', { multipart: { file: { name: 'sample.docx', mimeType: 'application/octet-stream', buffer }, edit: JSON.stringify(value) } });
  }
  expect((await edit({ kind: 'resize-image', objectId: 'i0', width: -1 })).status()).toBe(422);
  const result = await edit({ kind: 'replace-image', objectId: 'i0', imageBase64 });
  expect(result.ok()).toBe(true);
  const zip = await JSZip.loadAsync(await result.body());
  expect(await zip.file('customXml/preserved.xml').async('string')).toContain('Untouched package part');
  const inspection = await request.post('/api/documents/inspect', { data: await result.body() });
  const data = await inspection.json();
  expect(data.pages[0].images[0].partName).not.toBe(data.pages[0].images[1].partName);
  expect((await request.post('/api/documents/pdf?pages=0', { data: buffer })).status()).toBe(422);
});


test('image picker draws, creates QR codes and replaces a selected image', async ({ page }) => {
  await open(page);
  await page.getByRole('button', { name: 'Add image', exact: true }).click();
  let dialog = page.getByRole('dialog', { name: 'Add image', exact: true });
  await dialog.getByRole('button', { name: 'Draw image', exact: true }).click();
  const drawing = page.getByRole('dialog', { name: 'Draw an image', exact: true });
  const canvas = await drawing.getByLabel('Drawing canvas').boundingBox();
  await page.mouse.move(canvas.x + 20, canvas.y + 20);
  await page.mouse.down();
  await page.mouse.move(canvas.x + 100, canvas.y + 40, { steps: 5 });
  await page.mouse.up();
  await drawing.getByRole('button', { name: 'Use drawing', exact: true }).click();
  await dialog.getByRole('button', { name: 'Insert image', exact: true }).click();
  await expect(dialog).toHaveCount(0);
  await expect(page.locator('.canvas-host:visible [data-object-id^="i"]')).toHaveCount(3);
  await page.getByRole('button', { name: 'Select objects', exact: true }).click();
  await page.locator('.canvas-host:visible [data-object-id="i0"]').click();
  await page.getByRole('button', { name: 'Replace image', exact: true }).click();
  dialog = page.getByRole('dialog', { name: 'Replace image', exact: true });
  await dialog.getByRole('button', { name: 'Generate QR code', exact: true }).click();
  const qr = page.getByRole('dialog', { name: 'Generate QR code', exact: true });
  await qr.getByLabel('QR content').fill('https://example.com/docx');
  await qr.getByRole('button', { name: 'Use QR code', exact: true }).click();
  await dialog.getByRole('button', { name: 'Replace image', exact: true }).click();
  await expect(dialog).toHaveCount(0);
  await expect(page.locator('.canvas-host:visible [data-object-id^="i"]')).toHaveCount(3);
});


test('native canvas crops and rotates images with matching selection bounds', async ({ page }) => {
  await page.goto('/');
  const pixels = await page.evaluate(async () => {
    const { createDocxPreview } = await import('/src/renderers/docxRenderer.ts');
    const source = document.createElement('canvas');
    source.width = 100;
    source.height = 50;
    const paint = source.getContext('2d');
    paint.fillStyle = '#ff0000';
    paint.fillRect(0, 0, 50, 50);
    paint.fillStyle = '#0000ff';
    paint.fillRect(50, 0, 50, 50);
    const preview = await createDocxPreview({ width: 200, height: 200, warnings: [], media: { test: source.toDataURL() }, pages: [{ number: 1, id: 'test', title: '', textObjects: [], runs: [], rules: [], images: [{ id: 'i0', name: 'Synthetic', partName: 'test', x: 80, y: 80, width: 80, height: 40, rotation: 90, cropLeft: .5 }] }] });
    const canvas = document.createElement('canvas');
    canvas.width = 200;
    canvas.height = 200;
    const layer = document.createElement('div');
    await preview.render(canvas, 1, new AbortController().signal, layer);
    const context = canvas.getContext('2d');
    return { rotated: Array.from(context.getImageData(120, 65, 1, 1).data), outside: Array.from(context.getImageData(85, 100, 1, 1).data), transform: layer.querySelector('[data-object-id="i0"]').style.transform };
  });
  expect(pixels.rotated).toEqual([0, 0, 255, 255]);
  expect(pixels.outside).toEqual([255, 255, 255, 255]);
  expect(pixels.transform).toBe('rotate(90deg)');
});

test('refresh TOC uses current heading pages and survives undo and download', async ({ page }) => {
  const zip = await JSZip.loadAsync(await readFile(fixture));
  const xml = await zip.file('word/document.xml').async('string');
  zip.file('word/document.xml', xml.replace(/<w:body>[\s\S]*<\/w:body>/, `<w:body>
    <w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Stale contents</w:t></w:r></w:p></w:sdtContent></w:sdt>
    <w:p><w:pPr><w:outlineLvl w:val="0"/><w:pageBreakBefore/></w:pPr><w:r><w:t>Current heading</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body>`));
  await page.goto('/');
  await page.getByLabel('Open DOCX', { exact: true }).setInputFiles({ name: 'contents.docx', mimeType: 'application/octet-stream', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.canvas-host:visible .text-layer')).toContainText('Stale contents');
  await menu(page, 'Refresh table of contents');
  await expect(page.locator('.canvas-host:visible .text-layer')).toContainText('Current heading');
  await expect(page.locator('.canvas-host:visible .text-layer')).not.toContainText('Stale contents');
  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download DOCX', exact: true }).click();
  const output = await JSZip.loadAsync(await readFile(await (await download).path()));
  const updated = await output.file('word/document.xml').async('string');
  expect(updated).toContain('PAGEREF');
  expect(updated).toMatch(/<w:t>2<\/w:t>/);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.locator('.canvas-host:visible .text-layer')).toContainText('Stale contents');
});
