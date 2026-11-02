import { test, expect } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { toolbarAction } from './toolbar.js';

test('export editable Word, change options and download without changing PDF', async ({ page }, testInfo) => {
  await page.goto('/');
  await page.getByLabel('Open PDF', { exact: true }).setInputFiles(pdfFile('word-source.pdf', ['Editable first page', 'Second page']));
  await (await toolbarAction(page, 'Export Word')).click();
  const dialog = page.getByRole('dialog', { name: 'Export editable Word' });
  await dialog.getByRole('button', { name: 'Convert to Word', exact: true }).click();
  const result = dialog.getByRole('region', { name: 'Word export result' });
  await expect(result).toContainText('2 source pages converted');
  await expect(result).toContainText('2 paragraphs');
  await dialog.locator('summary').filter({ hasText: 'Conversion notes' }).click();
  await expect(result).toContainText('reflowable');
  await dialog.getByLabel('Pages', { exact: true }).selectOption('current');
  await expect(result).toHaveCount(0);
  await dialog.getByLabel('Include pictures').uncheck();
  await dialog.getByLabel('Keep source page breaks').uncheck();
  await dialog.getByRole('button', { name: 'Convert to Word', exact: true }).click();
  await expect(result).toContainText('1 source page converted');
  const downloadEvent = page.waitForEvent('download');
  await dialog.getByRole('button', { name: 'Download DOCX' }).click();
  const download = await downloadEvent;
  expect(download.suggestedFilename()).toBe('word-source.docx');
  const bytes = await readFile(await download.path());
  const zip = await JSZip.loadAsync(bytes);
  const xml = await zip.file('word/document.xml').async('string');
  expect(xml).toContain('Editable');
  expect(xml).not.toContain('Second');
  await dialog.screenshot({ path: testInfo.outputPath('word-export.png') });
  await dialog.getByRole('button', { name: 'Close Word export' }).click();
  await expect(dialog).toHaveCount(0);
  await (await toolbarAction(page, 'Export Word')).click();
  await page.getByRole('button', { name: 'Convert to Word', exact: true }).click();
  await expect(page.getByRole('region', { name: 'Word export result' })).toContainText('2 source pages converted');
});

test('Word export API validates options and request content', async ({ request }) => {
  const file = pdfFile('input.pdf', ['A page']);
  for (const options of [{ pageNumbers: [0] }, { pageNumbers: [1, 1] }, { unknown: true }]) {
    const response = await request.post(`/api/documents/docx?options=${encodeURIComponent(JSON.stringify(options))}`, {
      headers: { 'Content-Type': 'application/pdf' }, data: file.buffer,
    });
    expect(response.status()).toBe(422);
  }
  const response = await request.post('/api/documents/docx', { data: 'not a PDF' });
  expect(response.status()).toBe(415);
});

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

