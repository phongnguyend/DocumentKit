import { test, expect } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { toolbarAction } from './toolbar.js';

function pdfFile(name, texts, tagged = false, rotation = 0) {
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Count ' + texts.length + ' /Kids [' + texts.map((_, i) => (5 + i * 2) + ' 0 R').join(' ') + '] >>',
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
  ];
  for (const text of texts) {
    let content = 'BT /F1 12 Tf 20 250 Td ' + text.split(' ').map((word, index) => '(' + (index ? ' ' : '') + word + ') Tj').join(' ') + ' ET';
    if (tagged) {
      content = '/P << /MCID 0 >> BDC ' + content + ' EMC';
    }
    const stream = objects.length + 1;
    objects.push('<< /Length ' + content.length + ' >>\nstream\n' + content + '\nendstream');
    objects.push('<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Rotate ' + rotation + ' /Resources << /Font << /F1 3 0 R >> >> /Contents ' + stream + ' 0 R >>');
  }
  if (tagged) {
    const root = objects.length + 1;
    const elements = texts.map((_, i) => root + i + 1);
    objects[0] = '<< /Type /Catalog /Pages 2 0 R /MarkInfo << /Marked true >> /StructTreeRoot ' + root + ' 0 R >>';
    objects.push('<< /Type /StructTreeRoot /K [' + elements.map(id => id + ' 0 R').join(' ') + '] /ParentTree << /Nums [' + elements.map((id, i) => i + ' [' + id + ' 0 R]').join(' ') + '] >> >>');
    for (let i = 0; i < texts.length; i++) {
      objects[4 + i * 2] = objects[4 + i * 2].slice(0, -2) + ' /StructParents ' + i + ' >>';
      objects.push('<< /Type /StructElem /S /P /P ' + root + ' 0 R /Pg ' + (5 + i * 2) + ' 0 R /K 0 >>');
    }
  }
  let pdf = '%PDF-1.7\n';
  const offsets = [0];
  objects.forEach((value, index) => {
    offsets.push(Buffer.byteLength(pdf));
    pdf += (index + 1) + ' 0 obj\n' + value + '\nendobj\n';
  });
  const xref = Buffer.byteLength(pdf);
  pdf += 'xref\n0 ' + (objects.length + 1) + '\n0000000000 65535 f \n';
  pdf += offsets.slice(1).map(offset => String(offset).padStart(10, '0') + ' 00000 n \n').join('');
  pdf += 'trailer\n<< /Size ' + (objects.length + 1) + ' /Root 1 0 R >>\nstartxref\n' + xref + '\n%%EOF\n';
  return { name, mimeType: 'application/pdf', buffer: Buffer.from(pdf) };
}

test('export editable PowerPoint, change options, download and retain source PDF', async ({ page, request }, testInfo) => {
  const file = pdfFile('editable.pdf', ['Editable first slide', 'Second slide']);
  await page.goto('/');
  await page.getByLabel('Open PDF', { exact: true }).setInputFiles(file);
  await (await toolbarAction(page, 'Export PowerPoint')).click();
  const dialog = page.getByRole('dialog', { name: 'Export editable PowerPoint' });
  await dialog.getByRole('button', { name: 'Preview', exact: true }).click();
  const result = dialog.getByRole('region', { name: 'PowerPoint export result' });
  await expect(result).toContainText('2 slides ready');
  await expect(result).toContainText('2 editable text boxes');
  const preview = dialog.getByRole('region', { name: 'PowerPoint preview' });
  const slide = preview.getByRole('img', { name: 'PowerPoint slide 1' });
  await expect(slide).toBeVisible();
  await preview.getByRole('button', { name: 'Selectable text', exact: true }).click();
  const selectable = preview.locator('.pptx-preview-text');
  await expect(selectable).toContainText('Editable first slide');
  await selectable.locator('span').first().dblclick();
  await expect.poll(() => page.evaluate(() => window.getSelection()?.toString())).toBe('Editable');
  const layerBounds = await selectable.boundingBox();
  const canvasBounds = await slide.boundingBox();
  expect(Math.abs(layerBounds.width - canvasBounds.width)).toBeLessThan(1);
  expect(Math.abs(layerBounds.height - canvasBounds.height)).toBeLessThan(1);
  await preview.getByRole('button', { name: 'Canvas', exact: true }).click();
  await expect(selectable).toHaveCount(0);

  expect(await slide.evaluate(canvas => {
    const pixels = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
    let ink = 0;
    for (let i = 0; i < pixels.length; i += 4) {
      if (pixels[i] < 100 && pixels[i + 1] < 100 && pixels[i + 2] < 100) {
        ink++;
      }
    }
    return ink;
  })).toBeGreaterThan(50);
  await preview.getByRole('button', { name: 'Next preview slide' }).click();
  await expect(preview.getByRole('img', { name: 'PowerPoint slide 2' })).toBeVisible();
  await preview.getByRole('button', { name: 'Previous preview slide' }).click();
  await expect(slide).toBeVisible();

  await dialog.getByRole('button', { name: 'Hide preview' }).click();
  await page.setViewportSize({ width: 1280, height: 720 });
  await expect.poll(() => dialog.locator('.pptx-export-body').evaluate(element => element.scrollHeight - element.clientHeight)).toBeLessThanOrEqual(1);
  await expect.poll(() => dialog.evaluate(element => element.scrollHeight - element.clientHeight)).toBeLessThanOrEqual(1);
  await dialog.screenshot({ path: testInfo.outputPath('pptx-settings-720.png') });
  await dialog.getByRole('button', { name: 'Preview', exact: true }).click();
  await expect(preview.getByRole('img', { name: 'PowerPoint slide 1' })).toBeVisible();
  await dialog.screenshot({ path: testInfo.outputPath('pptx-export.png') });
  await dialog.getByText('Conversion notes').click();
  await expect(result).toContainText('Fonts are referenced by name');
  await dialog.getByLabel('Pages', { exact: true }).selectOption('current');
  await expect(result).toHaveCount(0);
  await expect(preview).toHaveCount(0);
  await dialog.getByLabel('Slide size').selectOption('wide');
  await dialog.getByLabel('Include pictures').uncheck();
  await dialog.getByLabel('Include vector shapes').uncheck();
  const exported = page.waitForResponse(response => response.url().includes('/api/documents/pptx?'));
  await dialog.getByRole('button', { name: 'Convert to PowerPoint' }).click();
  const data = await (await exported).json();
  expect(data.slideCount).toBe(1);
  expect(data.textBoxCount).toBe(1);
  await expect(result).toContainText('1 slide ready');
  await dialog.getByRole('button', { name: 'Preview', exact: true }).click();
  await expect(preview.getByRole('img', { name: 'PowerPoint slide 1' })).toBeVisible();
  await expect(preview.getByRole('button', { name: 'Next preview slide' })).toBeDisabled();
  const downloadPromise = page.waitForEvent('download');
  await dialog.getByRole('button', { name: 'Download PPTX' }).click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe('editable.pptx');
  const bytes = await readFile(await download.path());
  expect(bytes.subarray(0, 2).toString()).toBe('PK');
  expect(bytes.equals(Buffer.from(data.base64, 'base64'))).toBe(true);
  await dialog.getByRole('button', { name: 'Close PowerPoint export' }).click();
  await expect(page.getByRole('tab')).toHaveAccessibleName('editable.pdf');
  const invalid = await request.post('/api/documents/pptx?options=' + encodeURIComponent(JSON.stringify({ pageNumbers: [3] })), { data: file.buffer, headers: { 'Content-Type': 'application/pdf' } });
  expect(invalid.status()).toBe(422);
});
