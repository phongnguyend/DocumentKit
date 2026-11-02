import { test, expect } from '@playwright/test';
import { readFile } from 'node:fs/promises';

const fixture = new URL('../../PptxObjectKit.Tests/Fixtures/sample.pptx', import.meta.url);

test('open, select text, edit, reorder, undo, Markdown, delete and download', async ({ page, request }, testInfo) => {
  await page.goto('/');
  await page.getByLabel('Open PPTX', { exact: true }).setInputFiles({ name: 'sample.pptx', mimeType: 'application/octet-stream', buffer: await readFile(fixture) });
  await expect(page.getByRole('img', { name: 'Slide 1', exact: true })).toBeVisible();
  await expect(page.locator('.text-layer')).toContainText('PowerPoint to PDF test slide');
  for (const viewport of [{ width: 1440, height: 900 }, { width: 1280, height: 600 }]) {
    await page.setViewportSize(viewport);
    await expect.poll(() => page.locator('.slide-workspace').evaluate(element => element.scrollHeight <= element.clientHeight)).toBe(true);
    await expect.poll(() => page.locator('.canvas-host canvas').evaluate(canvas => {
      const bounds = canvas.getBoundingClientRect();
      const host = canvas.closest('.canvas-host').getBoundingClientRect();
      return bounds.height > 0 && bounds.bottom <= host.bottom + 1 && bounds.right <= host.right + 1;
    })).toBe(true);
  }
  await page.setViewportSize({ width: 1440, height: 900 });
  const firstThumbnail = page.locator('.slides-panel li').first().locator('.slide-thumbnail img');
  await expect(firstThumbnail).toBeVisible();
  const originalThumbnail = await firstThumbnail.getAttribute('src');
  await expect(page.getByRole('button', { name: 'Select text', exact: true })).toHaveAttribute('aria-pressed', 'true');
  await page.getByRole('button', { name: 'Select objects', exact: true }).click();
  await page.locator('.text-layer span[data-object-id]').first().click();
  await expect(page.getByLabel('Text', { exact: true })).toHaveValue('PowerPoint to PDF test slide');
  await expect(page.locator('.selected-object')).toHaveCount(1);
  await expect(page.locator('.selected-object')).toBeVisible();
  expect(await page.locator('.selected-object').evaluate(hit => {
    const rect = hit.getBoundingClientRect();
    const fragments = [...hit.parentElement.querySelectorAll('[data-text-object-id]')].filter(fragment => fragment.dataset.textObjectId === hit.dataset.objectId);
    return fragments.length > 1 && fragments.every(fragment => {
      const bounds = fragment.getBoundingClientRect();
      return bounds.left >= rect.left - 1 && bounds.right <= rect.right + 1 && bounds.top >= rect.top - 1 && bounds.bottom <= rect.bottom + 1;
    });
  })).toBe(true);
  await page.getByRole('button', { name: 'Select text', exact: true }).click();
  await page.getByLabel('Text', { exact: true }).fill('Updated presentation title');
  await page.getByRole('button', { name: 'Save text' }).click();
  await expect(page.locator('.text-layer')).toContainText('Updated presentation title');
  await expect(firstThumbnail).toBeVisible();
  await expect(firstThumbnail).not.toHaveAttribute('src', originalThumbnail);
  await page.locator('summary').filter({ hasText: /^Slides$/ }).click();
  await page.getByRole('button', { name: 'Move later' }).click();
  await expect(page.locator('.canvas-toolbar')).toContainText('Slide 2 of 2');
  await expect(page.locator('.slides-panel li').last()).toContainText('Updated presentation title');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.locator('.slides-panel li').first()).toContainText('Updated presentation title');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(page.locator('.slides-panel li').last()).toContainText('Updated presentation title');
  expect(await page.evaluate(() => document.documentElement.scrollHeight <= window.innerHeight)).toBe(true);
  await expect(page.getByRole('toolbar', { name: 'Slide navigation' })).toBeInViewport();
  await page.screenshot({ path: testInfo.outputPath('playground.png'), fullPage: true });
  await page.locator('summary').filter({ hasText: /^Document$/ }).click();
  await page.getByRole('button', { name: 'View Markdown' }).click();
  const dialog = page.getByRole('dialog', { name: 'View Markdown' });
  await expect(dialog.locator('article')).toContainText('Updated presentation title');
  await dialog.getByRole('button', { name: 'Source', exact: true }).click();
  await expect(dialog.locator('pre')).toContainText('# Slide 1: Second slide');
  await dialog.getByLabel('Speaker notes').uncheck();
  await dialog.getByRole('button', { name: 'Reapply options' }).click();
  await expect(dialog.getByRole('button', { name: 'Reapply options' })).toBeEnabled();
  await dialog.getByRole('button', { name: 'Close Markdown' }).click();
  await page.locator('summary').filter({ hasText: /^Slides$/ }).click();
  await page.getByRole('button', { name: 'Delete slide', exact: true }).click();
  await page.getByRole('dialog').getByRole('button', { name: 'Delete slide', exact: true }).click();
  await expect(page.locator('.slides-panel li')).toHaveCount(1);
  const downloaded = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download PPTX' }).click();
  const file = await downloaded;
  const bytes = await readFile(await file.path());
  const response = await request.post('/api/documents/inspect', { data: bytes });
  expect(response.ok()).toBe(true);
  expect((await response.json()).slides[0].title).toBe('Updated presentation title');
  await page.setViewportSize({ width: 390, height: 844 });
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBe(390);
  expect(await page.evaluate(() => document.documentElement.scrollHeight)).toBe(844);
  await expect(page.getByRole('toolbar', { name: 'Slide navigation' })).toBeInViewport();
});

test('reject invalid packages without losing the open presentation', async ({ page }) => {
  await page.goto('/');
  await page.getByLabel('Open PPTX', { exact: true }).setInputFiles({ name: 'sample.pptx', mimeType: 'application/octet-stream', buffer: await readFile(fixture) });
  await expect(page.locator('.text-layer')).toContainText('PowerPoint to PDF test slide');
  await page.getByLabel('Open PPTX', { exact: true }).setInputFiles({ name: 'broken.pptx', mimeType: 'application/octet-stream', buffer: Buffer.from('not a zip') });
  await expect(page.getByRole('alert')).toContainText('Cannot process');
  await expect(page.locator('.text-layer')).toContainText('PowerPoint to PDF test slide');
});


test('preview resolves layout geometry, theme fonts/colors and wraps centered text', async ({ page }) => {
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(await readFile(fixture));
  const ns = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
  zip.file('ppt/slides/slide1.xml', `<p:sld ${ns}><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>Wrapping title across several lines</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`);
  const relation = (type, target) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/></Relationships>`;
  zip.file('ppt/slides/_rels/slide1.xml.rels', relation('slideLayout', '../slideLayouts/layout.xml'));
  zip.file('ppt/slideLayouts/layout.xml', `<p:sldLayout ${ns}><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1270000" y="1270000"/><a:ext cx="2540000" cy="2540000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="3200"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></a:lstStyle></p:txBody></p:sp></p:spTree></p:cSld></p:sldLayout>`);
  zip.file('ppt/slideLayouts/_rels/layout.xml.rels', relation('slideMaster', '../slideMasters/master.xml'));
  zip.file('ppt/slideMasters/master.xml', `<p:sldMaster ${ns}><p:cSld><p:spTree/></p:cSld></p:sldMaster>`);
  zip.file('ppt/slideMasters/_rels/master.xml.rels', relation('theme', '../theme/theme.xml'));
  zip.file('ppt/theme/theme.xml', `<a:theme ${ns}><a:themeElements><a:clrScheme><a:accent1><a:srgbClr val="FF0000"/></a:accent1></a:clrScheme><a:fontScheme><a:majorFont><a:latin typeface="Arial"/></a:majorFont></a:fontScheme></a:themeElements></a:theme>`);
  await page.goto('/');
  await page.getByLabel('Open PPTX', { exact: true }).setInputFiles({ name: 'theme.pptx', mimeType: 'application/octet-stream', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.text-layer')).toContainText('Wrapping title across several lines');
  const rendered = await page.evaluate(() => {
    const spans = [...document.querySelectorAll('.text-layer span')].filter(span => span.textContent.length > 0);
    const canvas = document.querySelector('.canvas-host canvas');
    const pixels = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
    let red = 0;
    for (let i = 0; i < pixels.length; i += 4) {
      if (pixels[i] > 200 && pixels[i + 1] < 80 && pixels[i + 2] < 80) {
        red++;
      }
    }
    return { fonts: spans.map(span => span.style.font), rows: new Set(spans.map(span => new DOMMatrix(span.style.transform).f)).size, red };
  });
  expect(rendered.fonts.length).toBeGreaterThan(0);
  expect(rendered.fonts.every(font => font.includes('32px') && font.includes('Arial'))).toBe(true);
  expect(rendered.rows).toBeGreaterThan(1);
  expect(rendered.red).toBeGreaterThan(100);
});


test('agenda keeps explicit gradients, list markers and paragraph spacing', async ({ page }) => {
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(await readFile(fixture));
  const ns = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
  const paragraph = (text, bullet = '<a:buAutoNum type="arabicPeriod"/>') => `<a:p><a:pPr marL="254000" indent="-254000"><a:spcBef><a:spcPts val="1000"/></a:spcBef>${bullet}</a:pPr><a:r><a:rPr sz="2400"><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="00AA00"/></a:gs><a:gs pos="100000"><a:srgbClr val="00FF00"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill></a:rPr><a:t>${text}</a:t></a:r></a:p>`;
  zip.file('ppt/slides/slide1.xml', `<p:sld ${ns}><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Agenda"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="1270000" y="1270000"/><a:ext cx="5080000" cy="3810000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr><a:normAutofit lnSpcReduction="10000"/></a:bodyPr><a:lstStyle><a:lvl1pPr><a:defRPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>${paragraph('First')}${paragraph('Second')}${paragraph('Child', '<a:buFont typeface="Wingdings"/><a:buChar char="&#167;"/>')}</p:txBody></p:sp></p:spTree></p:cSld></p:sld>`);
  await page.goto('/');
  await page.getByLabel('Open PPTX', { exact: true }).setInputFiles({ name: 'agenda.pptx', mimeType: 'application/octet-stream', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.text-layer')).toContainText('1.First2.Second\u25aaChild');
  const result = await page.evaluate(() => {
    const spans = [...document.querySelectorAll('.text-layer span')].filter(span => span.textContent.length > 0);
    const position = text => new DOMMatrix(spans.find(span => span.textContent === text).style.transform).f;
    const canvas = document.querySelector('.canvas-host canvas');
    const data = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
    let green = 0;
    for (let i = 0; i < data.length; i += 4) {
      if (data[i] < 80 && data[i + 1] > 130 && data[i + 2] < 80) {
        green++;
      }
    }
    return { green, gap: (position('Second') - position('First')) / (canvas.width / 720) };
  });
  expect(result.green).toBeGreaterThan(100);
  expect(result.gap).toBeCloseTo(35.9, 0);
});


test('renders group fills and cached SmartArt with readable flipped text', async ({ page }) => {
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(await readFile(fixture));
  const ns = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  zip.file('ppt/slides/slide1.xml', `<p:sld ${ns}><p:cSld><p:spTree>
    <p:grpSp><p:grpSpPr><a:xfrm><a:off x="1270000" y="1270000"/><a:ext cx="1270000" cy="1270000"/><a:chOff x="0" y="0"/><a:chExt cx="1270000" cy="1270000"/></a:xfrm><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></p:grpSpPr>
    <p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="1270000"/></a:xfrm><a:prstGeom prst="round2DiagRect"/><a:grpFill/></p:spPr></p:sp></p:grpSp>
    <p:graphicFrame><p:xfrm><a:off x="2540000" y="2540000"/><a:ext cx="3810000" cy="1270000"/></p:xfrm><a:graphic><a:graphicData><d:relIds xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="data"/></a:graphicData></a:graphic></p:graphicFrame>
    </p:spTree></p:cSld></p:sld>`);
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="data" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data.xml"/><Relationship Id="drawing" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing.xml"/></Relationships>`);
  zip.file('ppt/diagrams/data.xml', '<data><dataModelExt relId="drawing"/></data>');
  zip.file('ppt/diagrams/drawing.xml', `<p:drawing ${ns}><p:spTree><p:sp><p:spPr><a:xfrm flipH="1"><a:off x="0" y="0"/><a:ext cx="3810000" cy="1270000"/></a:xfrm><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="2400"/><a:t>Diagram text</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:drawing>`);
  await page.goto('/');
  await page.getByLabel('Open PPTX', { exact: true }).setInputFiles({ name: 'diagram.pptx', mimeType: 'application/octet-stream', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.text-layer')).toContainText('Diagram text');
  const result = await page.evaluate(() => {
    const span = [...document.querySelectorAll('.text-layer span')].find(span => span.textContent.length > 0);
    const canvas = document.querySelector('.canvas-host canvas');
    const context = canvas.getContext('2d');
    const pixel = (x, y) => [...context.getImageData(Math.round(x * canvas.width / 720), Math.round(y * canvas.width / 720), 1, 1).data];
    return { transform: new DOMMatrix(span.style.transform).a, inside: pixel(150, 150), roundedCorner: pixel(101, 101) };
  });
  expect(result.transform).toBeGreaterThan(0);
  expect(result.inside).toEqual([0, 0, 255, 255]);
  expect(result.roundedCorner).toEqual([255, 255, 255, 255]);
});


test('renders table cells and Wingdings symbols without replacing Latin text', async ({ page }) => {
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(await readFile(fixture));
  const ns = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
  const cell = text => `<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1800"><a:sym typeface="Wingdings"/></a:rPr><a:t>${text}</a:t></a:r></a:p></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></a:tcPr></a:tc>`;
  zip.file('ppt/slides/slide1.xml', `<p:sld ${ns}><p:cSld><p:spTree><p:graphicFrame><p:xfrm><a:off x="1270000" y="1270000"/><a:ext cx="5080000" cy="1270000"/></p:xfrm><a:graphic><a:graphicData><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="2540000"/><a:gridCol w="2540000"/></a:tblGrid><a:tr h="635000">${cell('Cohesion')}${cell('Coupling')}</a:tr><a:tr h="635000">${cell('&#61672;Unit testing')}${cell('Body')}</a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>`);
  await page.goto('/');
  await page.getByLabel('Open PPTX', { exact: true }).setInputFiles({ name: 'table.pptx', mimeType: 'application/octet-stream', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.text-layer')).toContainText('CohesionCoupling\u2794Unit testingBody');
  const result = await page.evaluate(() => {
    const canvas = document.querySelector('.canvas-host canvas');
    const scale = canvas.width / 720;
    const spans = [...document.querySelectorAll('.text-layer span')].filter(span => span.textContent.length > 0);
    const pos = value => new DOMMatrix(spans.find(span => span.textContent === value).style.transform);
    return { pixel: [...canvas.getContext('2d').getImageData(110 * scale, 140 * scale, 1, 1).data], dx: (pos('Coupling').e - pos('Cohesion').e) / scale, dy: (pos('Body').f - pos('Coupling').f) / scale };
  });
  expect(result.pixel).toEqual([0, 170, 0, 255]);
  expect(result.dx).toBeCloseTo(200);
  expect(result.dy).toBeCloseTo(50);
});


test('theme style references paint boxes and elbow arrow connectors', async ({ page }) => {
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(await readFile(fixture));
  const ns = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
  const rel = (type, target) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/></Relationships>`;
  zip.file('ppt/slides/_rels/slide1.xml.rels', rel('slideLayout', '../slideLayouts/layout.xml'));
  zip.file('ppt/slideLayouts/layout.xml', `<p:sldLayout ${ns}><p:cSld><p:spTree/></p:cSld></p:sldLayout>`);
  zip.file('ppt/slideLayouts/_rels/layout.xml.rels', rel('slideMaster', '../slideMasters/master.xml'));
  zip.file('ppt/slideMasters/master.xml', `<p:sldMaster ${ns}><p:cSld><p:spTree/></p:cSld></p:sldMaster>`);
  zip.file('ppt/slideMasters/_rels/master.xml.rels', rel('theme', '../theme/theme.xml'));
  zip.file('ppt/theme/theme.xml', `<a:theme ${ns}><a:themeElements><a:fmtScheme><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst></a:fmtScheme></a:themeElements></a:theme>`);
  zip.file('ppt/slides/slide1.xml', `<p:sld ${ns}><p:cSld><p:spTree><p:sp><p:spPr><a:xfrm><a:off x="1270000" y="1270000"/><a:ext cx="1270000" cy="1270000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:style><a:fillRef idx="1"><a:srgbClr val="FF0000"/></a:fillRef></p:style></p:sp><p:cxnSp><p:spPr><a:xfrm><a:off x="3810000" y="1270000"/><a:ext cx="1270000" cy="1270000"/></a:xfrm><a:prstGeom prst="bentConnector2"/><a:ln><a:tailEnd type="triangle"/></a:ln></p:spPr><p:style><a:lnRef idx="1"><a:srgbClr val="0000FF"/></a:lnRef></p:style></p:cxnSp></p:spTree></p:cSld></p:sld>`);
  await page.goto('/');
  await page.getByLabel('Open PPTX', { exact: true }).setInputFiles({ name: 'connector.pptx', mimeType: 'application/octet-stream', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.canvas-host canvas')).toBeVisible();
  const pixels = await page.evaluate(() => {
    const canvas = document.querySelector('.canvas-host canvas');
    const ctx = canvas.getContext('2d');
    const pixel = (x, y) => [...ctx.getImageData(Math.round(x * canvas.width / 720), Math.round(y * canvas.width / 720), 1, 1).data];
    return [pixel(150, 150), pixel(350, 100), pixel(350, 150), pixel(398, 195)];
  });
  expect(pixels[0]).toEqual([255, 0, 0, 255]);
  expect(pixels[1]).toEqual([0, 0, 255, 255]);
  expect(pixels[2]).toEqual([255, 255, 255, 255]);
  expect(pixels[3][2]).toBe(255);
  expect(pixels[3][0]).toBeLessThan(100);
});


test('long linked URLs wrap within the frame and preserve all selectable text', async ({ page }) => {
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(await readFile(fixture));
  const url = 'https://example.com/averylongpathwithoutspaces/Architecture';
  zip.file('ppt/slides/slide1.xml', `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:sp><p:spPr><a:xfrm><a:off x="1270000" y="1270000"/><a:ext cx="2540000" cy="2540000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1800"><a:hlinkClick r:id="link"/></a:rPr><a:t>${url}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`);
  await page.goto('/');
  await page.getByLabel('Open PPTX', { exact: true }).setInputFiles({ name: 'url.pptx', mimeType: 'application/octet-stream', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.text-layer')).toHaveText(url);
  const result = await page.evaluate(() => {
    const spans = [...document.querySelectorAll('.text-layer span')].filter(span => span.textContent.length > 0);
    const canvas = document.querySelector('.canvas-host canvas');
    const scale = canvas.width / 720;
    return { rows: new Set(spans.map(span => new DOMMatrix(span.style.transform).f)).size, right: Math.max(...spans.map(span => (new DOMMatrix(span.style.transform).e + parseFloat(span.style.width) * scale) / scale)) };
  });
  expect(result.rows).toBeGreaterThan(1);
  expect(result.right).toBeLessThanOrEqual(300);
});


test('wraps formatted words as units and renders the Wingdings sad-face symbol', async ({ page }) => {
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(await readFile(fixture));
  await page.goto('/');
  const width = await page.evaluate(() => {
    const ctx = document.createElement('canvas').getContext('2d');
    ctx.font = '18px Arial';
    return ctx.measureText('prefix too').width + 14.4 + .1;
  });
  const run = (text, extra = '') => `<a:r><a:rPr sz="1800">${extra}</a:rPr><a:t>${text}</a:t></a:r>`;
  zip.file('ppt/slides/slide1.xml', `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:spPr><a:xfrm><a:off x="1270000" y="1270000"/><a:ext cx="${Math.round(width * 12700)}" cy="3810000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p>${run('prefix too')}${run('.')}</a:p><a:p>${run('prefix ')}${run('A&#160;')}${run('B')}</a:p><a:p>${run('&#61516; Latin', '<a:sym typeface="Wingdings"/>')}</a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`);
  await page.getByLabel('Open PPTX', { exact: true }).setInputFiles({ name: 'runs.pptx', mimeType: 'application/octet-stream', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.text-layer')).toContainText('\u2639\ufe0e Latin');
  const rows = await page.evaluate(() => {
    const spans = [...document.querySelectorAll('.text-layer span')].filter(span => span.textContent.length > 0);
    const y = text => new DOMMatrix(spans.find(span => span.textContent === text).style.transform).f;
    return { word: y('too'), punctuation: y('.'), first: y('prefix'), nbsp: y('A\u00a0'), last: y('B') };
  });
  expect(rows.word).toBe(rows.punctuation);
  expect(rows.word).toBeGreaterThan(rows.first);
  expect(rows.nbsp).toBe(rows.last);
});


test('document tabs preserve slide selection, edits and history with safe closing', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('.application-header .brand')).toHaveCount(1);
  await expect(page.locator('.presentation-panel .brand')).toHaveCount(0);
  await expect(page.locator('.application-header').getByRole('button', { name: 'Open PPTX' })).toBeVisible();
  await expect(page.locator('.document-toolbar')).toHaveCount(0);
  const buffer = await readFile(fixture);
  await page.getByLabel('Open PPTX', { exact: true }).setInputFiles([
    { name: 'first.pptx', mimeType: 'application/octet-stream', buffer },
    { name: 'second.pptx', mimeType: 'application/octet-stream', buffer },
  ]);
  await expect(page.getByRole('tab')).toHaveCount(2);
  await page.getByRole('tab', { name: 'first.pptx', exact: true }).press('End');
  await expect(page.getByRole('tab', { name: 'second.pptx', exact: true })).toHaveAttribute('aria-selected', 'true');
  await page.getByRole('tab', { name: 'second.pptx', exact: true }).press('Home');
  await page.locator('.document-tab').nth(1).dragTo(page.locator('.document-tab').nth(0));
  await expect(page.getByRole('tab').first()).toHaveText('second.pptx');
  await page.getByRole('tab', { name: 'first.pptx', exact: true }).click();
  const panel = page.getByRole('tabpanel');
  await panel.locator('.inspector .object-list button').first().click();
  await panel.getByRole('textbox', { name: 'Text', exact: true }).fill('First document edited');
  await panel.getByRole('button', { name: 'Save text' }).click();
  await expect(panel.locator('.text-layer')).toContainText('First document edited');
  await panel.locator('.slides-panel li button').nth(1).click();
  await expect(panel.locator('.canvas-toolbar')).toContainText('Slide 2 of 2');
  await page.getByRole('tab', { name: 'second.pptx', exact: true }).click();
  await expect(panel.locator('.text-layer')).toContainText('PowerPoint to PDF test slide');
  await expect(panel.getByRole('button', { name: 'Undo', exact: true })).toBeDisabled();
  await page.getByRole('tab', { name: /first.pptx/ }).click();
  await expect(panel.locator('.canvas-toolbar')).toContainText('Slide 2 of 2');
  await panel.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(panel.locator('.text-layer')).toContainText('PowerPoint to PDF test slide');
  await panel.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(panel.locator('.text-layer')).toContainText('First document edited');
  await page.screenshot({ path: 'test-results/document-tabs.png' });
  await page.getByRole('button', { name: 'Close first.pptx', exact: true }).click();
  await expect(page.getByRole('dialog')).toContainText('have not been downloaded');
  await page.getByRole('button', { name: 'Cancel', exact: true }).click();
  await expect(page.getByRole('tab')).toHaveCount(2);
  await page.getByRole('button', { name: 'Close first.pptx', exact: true }).click();
  await page.getByRole('button', { name: 'Discard and close' }).click();
  await expect(page.getByRole('tab')).toHaveCount(1);
  await expect(panel.locator('.text-layer')).toContainText('PowerPoint to PDF test slide');
  await page.getByRole('button', { name: 'Close second.pptx', exact: true }).click();
  await expect(page.getByRole('tab', { name: 'New presentation' })).toBeVisible();
  await page.getByRole('button', { name: 'New tab', exact: true }).click();
  await expect(page.getByRole('tab')).toHaveCount(2);
});


test('manage image instances with insert, replace, resize, download, delete and undo', async ({ page, request }) => {
  await page.goto('/');
  await page.getByLabel('Open PPTX', { exact: true }).setInputFiles({ name: 'images.pptx', mimeType: 'application/octet-stream', buffer: await readFile(fixture) });
  await expect(page.getByRole('img', { name: 'Slide 1', exact: true })).toBeVisible();
  const image = await page.evaluate(() => {
    const canvas = document.createElement('canvas');
    canvas.width = 40;
    canvas.height = 30;
    const context = canvas.getContext('2d');
    context.fillStyle = 'red';
    context.fillRect(0, 0, 40, 30);
    return canvas.toDataURL().split(',')[1];
  });
  const upload = { name: 'red.png', mimeType: 'image/png', buffer: Buffer.from(image, 'base64') };
  await page.getByRole('button', { name: 'Add image', exact: true }).click();
  await page.getByRole('dialog', { name: 'Add image', exact: true }).getByLabel('Image file', { exact: true }).setInputFiles(upload);
  await page.getByRole('button', { name: 'Insert image', exact: true }).click();
  await expect(page.locator('.text-layer [data-kind="image"]')).toHaveCount(1);
  await page.getByRole('button', { name: 'Select objects', exact: true }).click();
  await page.locator('.text-layer [data-kind="image"]').click({ position: { x: 5, y: 5 } });
  await expect(page.getByRole('button', { name: 'Download image', exact: true })).toBeVisible();
  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download image', exact: true }).click();
  expect(await readFile(await (await download).path())).toEqual(upload.buffer);
  const originalX = Number(await page.getByLabel('Image x', { exact: true }).inputValue());
  const transform = page.getByRole('group', { name: 'Move selected image' });
  const initial = await transform.boundingBox();
  await page.mouse.move(initial.x + initial.width / 2, initial.y + initial.height / 2);
  await page.mouse.down();
  await page.mouse.move(initial.x + initial.width / 2 + 35, initial.y + initial.height / 2 + 20, { steps: 5 });
  await page.keyboard.press('Escape');
  await page.mouse.up();
  await expect(page.getByLabel('Image x', { exact: true })).toHaveValue(String(originalX));
  await page.mouse.move(initial.x + initial.width / 2, initial.y + initial.height / 2);
  await page.mouse.down();
  await page.mouse.move(initial.x + initial.width / 2 + 35, initial.y + initial.height / 2 + 20, { steps: 5 });
  await page.mouse.up();
  await expect(page.locator('.image-editor')).toHaveCount(0);
  await page.locator('.inspector .object-list button').last().click();
  expect(Number(await page.getByLabel('Image x', { exact: true }).inputValue())).toBeGreaterThan(originalX);
  const originalWidth = Number(await page.getByLabel('Image width', { exact: true }).inputValue());
  const handle = await page.getByRole('button', { name: 'Resize image se', exact: true }).boundingBox();
  await page.mouse.move(handle.x + handle.width / 2, handle.y + handle.height / 2);
  await page.mouse.down();
  await page.mouse.move(handle.x + handle.width / 2 + 30, handle.y + handle.height / 2 + 20, { steps: 5 });
  await page.mouse.up();
  await expect(page.locator('.image-editor')).toHaveCount(0);
  await page.locator('.inspector .object-list button').last().click();
  expect(Number(await page.getByLabel('Image width', { exact: true }).inputValue())).toBeGreaterThan(originalWidth);
  await page.getByLabel('Image x', { exact: true }).fill('90');
  await page.getByLabel('Image width', { exact: true }).fill('240');
  await page.getByRole('button', { name: 'Save position and size' }).click();
  await expect(page.locator('.image-editor')).toHaveCount(0);
  await page.locator('.inspector .object-list button').last().click();
  await expect(page.getByLabel('Image x', { exact: true })).toHaveValue('90');
  await page.getByRole('button', { name: 'Replace image', exact: true }).click();
  await page.getByRole('dialog', { name: 'Replace image', exact: true }).getByLabel('Image file', { exact: true }).setInputFiles(upload);
  await page.getByRole('dialog', { name: 'Replace image', exact: true }).getByRole('button', { name: 'Replace image', exact: true }).click();
  await expect(page.locator('.image-editor')).toHaveCount(0);
  await page.locator('.inspector .object-list button').last().click();
  await expect(page.getByLabel('Image width', { exact: true })).toHaveValue('240');
  await page.getByRole('button', { name: 'Delete image', exact: true }).click();
  await page.getByRole('dialog').getByRole('button', { name: 'Delete image', exact: true }).click();
  await expect(page.locator('.text-layer [data-kind="image"]')).toHaveCount(0);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.locator('.text-layer [data-kind="image"]')).toHaveCount(1);
  const saved = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download PPTX', exact: true }).click();
  const response = await request.post('/api/documents/inspect', { data: await readFile(await (await saved).path()) });
  expect(response.ok()).toBe(true);
  expect((await response.json()).slides[0].images[0].width).toBe(240);
});


test('image picker supports drawing, QR replacement, drop and source sizing', async ({ page }, testInfo) => {
  await page.goto('/');
  await page.getByLabel('Open PPTX', { exact: true }).setInputFiles({ name: 'drawing.pptx', mimeType: 'application/octet-stream', buffer: await readFile(fixture) });
  await expect(page.getByRole('img', { name: 'Slide 1', exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Add image', exact: true }).click();
  await page.getByRole('button', { name: 'Draw image', exact: true }).click();
  const drawing = page.getByRole('dialog', { name: 'Draw an image', exact: true });
  const canvas = drawing.getByLabel('Drawing canvas');
  const box = await canvas.boundingBox();
  expect(box.width / box.height).toBeCloseTo(1.5, 1);
  await expect(drawing.getByRole('button', { name: 'Use drawing' })).toBeDisabled();
  await page.mouse.move(box.x + 20, box.y + 20);
  await page.mouse.down();
  await page.mouse.move(box.x + 70, box.y + 60, { steps: 6 });
  await page.mouse.up();
  await drawing.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(drawing.getByRole('button', { name: 'Use drawing' })).toBeDisabled();
  await canvas.click({ position: { x: 45, y: 40 } });
  await drawing.getByRole('button', { name: 'Use drawing' }).click();
  await expect(page.getByAltText('Selected image preview')).toBeVisible();
  await page.getByRole('button', { name: 'Insert image', exact: true }).click();
  await expect(page.locator('.text-layer [data-kind="image"]')).toHaveCount(1);
  await page.locator('.inspector .object-list button').last().click();
  await page.getByRole('button', { name: 'Replace image', exact: true }).click();
  await page.getByRole('button', { name: 'Generate QR code' }).click();
  await page.getByLabel('QR content').fill('https://example.com/pptx');
  await expect(page.getByAltText('QR code preview')).toBeVisible();
  await page.getByRole('button', { name: 'Use QR code' }).click();
  const replacement = page.getByRole('dialog', { name: 'Replace image', exact: true });
  await expect(replacement.getByLabel('Placement width')).toHaveValue('120');
  await expect(replacement.getByLabel('Placement height')).toHaveValue('120');
  await page.screenshot({ path: testInfo.outputPath('image-picker.png') });
  await replacement.getByRole('button', { name: 'Replace image', exact: true }).click();
  await expect(replacement).toHaveCount(0);
  await page.locator('.inspector .object-list button').last().click();
  await expect(page.getByLabel('Image width', { exact: true })).toHaveValue('120');
  await expect(page.getByLabel('Image height', { exact: true })).toHaveValue('120');
  await page.getByRole('button', { name: 'Add image', exact: true }).click();
  const transfer = await page.evaluateHandle(async () => {
    const canvas = document.createElement('canvas');
    canvas.width = 80;
    canvas.height = 60;
    const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/webp'));
    const data = new DataTransfer();
    data.items.add(new File([blob], 'dropped.webp', { type: 'image/webp' }));
    return data;
  });
  await page.locator('.image-drop-zone').dispatchEvent('drop', { dataTransfer: transfer });
  await transfer.dispose();
  await page.getByRole('button', { name: 'Use image size' }).click();
  await expect(page.getByLabel('Placement width')).toHaveValue('80');
  await expect(page.getByLabel('Placement height')).toHaveValue('60');
  await page.getByRole('button', { name: 'Insert image', exact: true }).click();
  await expect(page.locator('.text-layer [data-kind="image"]')).toHaveCount(2);
});


test('renders SVG-only pictures on slides and layouts with independent relationship IDs', async ({ page }) => {
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(await readFile(fixture));
  const ns = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"';
  const picture = x => `<p:pic><p:nvPicPr><p:cNvPr id="20" name="SVG icon"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="image"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${x}" y="1270000"/><a:ext cx="1270000" cy="1270000"/></a:xfrm></p:spPr></p:pic>`;
  const relationships = contents => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${contents}</Relationships>`;
  const relationship = (id, type, target) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`;
  zip.file('ppt/slides/slide1.xml', `<p:sld ${ns}><p:cSld><p:spTree>${picture(1270000)}</p:spTree></p:cSld></p:sld>`);
  zip.file('ppt/slides/_rels/slide1.xml.rels', relationships(relationship('image', 'image', '../media/red.svg') + relationship('layout', 'slideLayout', '../slideLayouts/svg.xml')));
  zip.file('ppt/slideLayouts/svg.xml', `<p:sldLayout ${ns}><p:cSld><p:spTree>${picture(3810000)}</p:spTree></p:cSld></p:sldLayout>`);
  zip.file('ppt/slideLayouts/_rels/svg.xml.rels', relationships(relationship('image', 'image', '../media/green.svg')));
  for (const [name, color] of [['red', '#ff0000'], ['green', '#00ff00']]) {
    zip.file(`ppt/media/${name}.svg`, `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 96 96"><rect width="96" height="96" fill="${color}"/></svg>`);
  }
  await page.goto('/');
  await page.getByLabel('Open PPTX', { exact: true }).setInputFiles({ name: 'svg-icons.pptx', mimeType: 'application/octet-stream', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.getByRole('img', { name: 'Slide 1', exact: true })).toBeVisible();
  const colors = await page.locator('.canvas-host canvas').evaluate(canvas => {
    const pixels = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
    let red = 0, green = 0;
    for (let i = 0; i < pixels.length; i += 4) {
      if (pixels[i] > 200 && pixels[i + 1] < 30) {
        red++;
      }
      if (pixels[i + 1] > 200 && pixels[i] < 30) {
        green++;
      }
    }
    return { red, green };
  });
  expect(colors.red).toBeGreaterThan(1000);
  expect(colors.green).toBeGreaterThan(1000);
  await expect(page.getByText(/Preview unavailable:/)).toHaveCount(0);
});


test('duplicate and insert blank slides, navigate, undo and save independently', async ({ page, request }) => {
  await page.goto('/');
  await page.getByLabel('Open PPTX', { exact: true }).setInputFiles({ name: 'insert.pptx', mimeType: 'application/octet-stream', buffer: await readFile(fixture) });
  await expect(page.getByRole('img', { name: 'Slide 1', exact: true })).toBeVisible();
  async function action(name) {
    await page.locator('summary').filter({ hasText: /^Slides$/ }).click();
    await page.getByRole('button', { name, exact: true }).click();
  }
  await action('Duplicate slide');
  await expect(page.getByRole('combobox', { name: 'Slide', exact: true })).toHaveValue('2');
  await expect(page.locator('.slides-panel li')).toHaveCount(3);
  await page.locator('.inspector .object-list button').first().click();
  await page.getByLabel('Text', { exact: true }).fill('Independent duplicate');
  await page.getByRole('button', { name: 'Save text' }).click();
  await expect(page.locator('.text-layer')).toContainText('Independent duplicate');
  await page.getByRole('combobox', { name: 'Slide', exact: true }).selectOption('1');
  await expect(page.locator('.text-layer')).toContainText('PowerPoint to PDF test slide');
  await action('Blank slide after current');
  await expect(page.getByRole('combobox', { name: 'Slide', exact: true })).toHaveValue('2');
  await expect(page.getByRole('img', { name: 'Slide 2', exact: true })).toBeVisible();
  await expect(page.locator('.text-layer span')).toHaveCount(0);
  expect(await page.locator('.canvas-host canvas').evaluate(canvas => {
    const rgba = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
    return rgba.every(value => value === 255);
  })).toBe(true);
  await action('Blank slide at end');
  await expect(page.getByRole('combobox', { name: 'Slide', exact: true })).toHaveValue('5');
  await expect(page.locator('.slides-panel li')).toHaveCount(5);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.locator('.slides-panel li')).toHaveCount(4);
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(page.locator('.slides-panel li')).toHaveCount(5);
  const saved = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download PPTX', exact: true }).click();
  const response = await request.post('/api/documents/inspect', { data: await readFile(await (await saved).path()) });
  expect(response.ok()).toBe(true);
  const slides = (await response.json()).slides;
  expect(slides).toHaveLength(5);
  expect(slides[2].title).toBe('Independent duplicate');
  expect(slides[4].textObjects).toHaveLength(0);
});


test('drag and resize text boxes with cancel, automatic save and undo', async ({ page, request }) => {
  await page.goto('/');
  await page.getByLabel('Open PPTX', { exact: true }).setInputFiles({ name: 'text-transform.pptx', mimeType: 'application/octet-stream', buffer: await readFile(fixture) });
  await expect(page.getByRole('img', { name: 'Slide 1', exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Select objects', exact: true }).click();
  await page.locator('.text-object-hit').first().click();
  const transform = page.getByRole('group', { name: 'Move selected text', exact: true });
  const original = await transform.boundingBox();
  await page.mouse.move(original.x + 20, original.y + 20);
  await page.mouse.down();
  await page.mouse.move(original.x + 60, original.y + 50, { steps: 4 });
  await page.keyboard.press('Escape');
  await page.mouse.up();
  expect((await transform.boundingBox()).x).toBeCloseTo(original.x, 0);
  await page.mouse.move(original.x + 20, original.y + 20);
  await page.mouse.down();
  await page.mouse.move(original.x + 60, original.y + 50, { steps: 4 });
  await page.mouse.up();
  await expect(transform).toHaveCount(0);
  await page.locator('.inspector .object-list button').first().click();
  expect((await transform.boundingBox()).x).toBeGreaterThan(original.x + 20);
  const beforeResize = await transform.boundingBox();
  const handle = await page.getByRole('button', { name: 'Resize text se', exact: true }).boundingBox();
  await page.mouse.move(handle.x + 5, handle.y + 5);
  await page.mouse.down();
  await page.mouse.move(handle.x - 75, handle.y + 25, { steps: 5 });
  await page.mouse.up();
  await expect(transform).toHaveCount(0);
  await page.locator('.inspector .object-list button').first().click();
  expect((await transform.boundingBox()).width).toBeLessThan(beforeResize.width - 40);
  await expect(page.locator('.text-layer')).toContainText('PowerPoint to PDF test slide');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await page.locator('.inspector .object-list button').first().click();
  expect((await transform.boundingBox()).width).toBeCloseTo(beforeResize.width, 0);
  const saved = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download PPTX', exact: true }).click();
  const response = await request.post('/api/documents/inspect', { data: await readFile(await (await saved).path()) });
  expect(response.ok()).toBe(true);
  expect((await response.json()).slides[0].textObjects[0].bounds.x).toBeGreaterThan(0);
});
