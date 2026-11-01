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
  await dialog.getByRole('button', { name: 'Convert to PowerPoint' }).click();
  const result = dialog.getByRole('region', { name: 'PowerPoint export result' });
  await expect(result).toContainText('2 slides ready');
  await expect(result).toContainText('2 editable text boxes');
  await dialog.screenshot({ path: testInfo.outputPath('pptx-export.png') });
  await dialog.getByText('Conversion notes').click();
  await expect(result).toContainText('Fonts are referenced by name');
  await dialog.getByLabel('Pages', { exact: true }).selectOption('current');
  await expect(result).toHaveCount(0);
  await dialog.getByLabel('Slide size').selectOption('wide');
  await dialog.getByLabel('Include pictures').uncheck();
  await dialog.getByLabel('Include vector shapes').uncheck();
  const exported = page.waitForResponse(response => response.url().includes('/api/documents/pptx?'));
  await dialog.getByRole('button', { name: 'Convert to PowerPoint' }).click();
  const data = await (await exported).json();
  expect(data.slideCount).toBe(1);
  expect(data.textBoxCount).toBe(1);
  await expect(result).toContainText('1 slide ready');
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
