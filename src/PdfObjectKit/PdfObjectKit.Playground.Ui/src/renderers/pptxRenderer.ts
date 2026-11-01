import type JSZip from 'jszip';

const emu = 12700;
const child = (node: Element | null, name: string) => Array.from(node?.children ?? []).find(item => item.localName === name) ?? null;
const descendants = (node: Element | Document, name: string) => Array.from(node.getElementsByTagNameNS('*', name));
const number = (node: Element | null, name: string, fallback = 0) => {
  const value = Number(node?.getAttribute(name) ?? fallback);
  if (!Number.isFinite(value)) {
    throw new Error('Invalid numeric value in PowerPoint preview.');
  }
  return value;
};
type Slide = { zip: JSZip; xml: Document; width: number; height: number; images: Map<string, string>; notes: string[] };
async function xml(zip: JSZip, path: string) {
  const entry = zip.file(path);
  if (!entry) {
    throw new Error(`PowerPoint preview is missing ${path}.`);
  }
  const document = new DOMParser().parseFromString(await entry.async('string'), 'application/xml');
  if (document.querySelector('parsererror')) {
    throw new Error('Invalid XML in PowerPoint preview.');
  }
  return document;
}

/** Opens only packages produced by the PDF-to-PPTX exporter. */
export async function createPptxPreview(base64: string) {
  const { default: Zip } = await import('jszip');
  const zip = await Zip.loadAsync(base64, { base64: true });
  const presentation = await xml(zip, 'ppt/presentation.xml');
  const size = descendants(presentation, 'sldSz')[0];
  const width = number(size, 'cx') / emu;
  const height = number(size, 'cy') / emu;
  const count = descendants(presentation, 'sldId').length;
  if (width <= 0 || height <= 0 || count === 0) {
    throw new Error('Invalid PowerPoint slide size or slide count.');
  }
  return {
    width, height, count,
    async render(canvas: HTMLCanvasElement, page: number, signal?: AbortSignal, textLayer?: HTMLDivElement) {
      if (page < 1 || page > count) {
        throw new Error('Invalid PowerPoint preview slide.');
      }
      const relationships = await xml(zip, `ppt/slides/_rels/slide${page}.xml.rels`);
      const images = new Map<string, string>();
      for (const relationship of descendants(relationships, 'Relationship')) {
        if (relationship.getAttribute('Type')?.endsWith('/image') && relationship.getAttribute('TargetMode') !== 'External') {
          const target = new URL(relationship.getAttribute('Target') ?? '', 'https://pptx.invalid/ppt/slides/slide1.xml');
          images.set(relationship.getAttribute('Id') ?? '', target.pathname.slice(1));
        }
      }
      const slide = { zip, xml: await xml(zip, `ppt/slides/slide${page}.xml`), width, height, images, notes: [] };
      await renderSlide(canvas, slide, signal, textLayer);
    },
  };
}

function color(node: Element | null): string {
  const rgb = node?.getAttribute('val') ?? '000000';
  if (!/^[0-9a-f]{6}$/i.test(rgb)) {
    return '#000000';
  }
  return `rgba(${parseInt(rgb.slice(0, 2), 16)},${parseInt(rgb.slice(2, 4), 16)},${parseInt(rgb.slice(4, 6), 16)},${number(child(node, 'alpha'), 'val', 100000) / 100000})`;
}

function paint(context: CanvasRenderingContext2D, node: Element | null, width: number, height: number): string | CanvasGradient | null {
  if (child(node, 'noFill')) {
    return null;
  }
  const solid = child(node, 'solidFill');
  if (solid) {
    return color(child(solid, 'srgbClr'));
  }
  const fill = child(node, 'gradFill');
  if (!fill) {
    return null;
  }
  const angle = number(child(fill, 'lin'), 'ang') / 60000 * Math.PI / 180;
  const dx = Math.cos(angle), dy = Math.sin(angle);
  const length = Math.abs(width * dx) + Math.abs(height * dy);
  const gradient = context.createLinearGradient(width / 2 - dx * length / 2, height / 2 - dy * length / 2, width / 2 + dx * length / 2, height / 2 + dy * length / 2);
  for (const stop of descendants(fill, 'gs')) {
    gradient.addColorStop(Math.max(0, Math.min(1, number(stop, 'pos') / 100000)), color(child(stop, 'srgbClr')));
  }
  return gradient;
}

function text(context: CanvasRenderingContext2D, body: Element, width: number, height: number, textLayer?: HTMLDivElement) {
  let y = number(child(body, 'bodyPr'), 'tIns') / emu;
  for (const paragraph of Array.from(body.children).filter(item => item.localName === 'p')) {
    const runs = Array.from(paragraph.children).filter(item => item.localName === 'r');
    const size = Math.max(1, ...runs.map(run => number(child(run, 'rPr'), 'sz', 1200) / 100));
    let x = number(child(body, 'bodyPr'), 'lIns') / emu;
    y += size * .85;
    for (const run of runs) {
      const properties = child(run, 'rPr');
      const fontSize = number(properties, 'sz', 1200) / 100;
      const family = child(properties, 'latin')?.getAttribute('typeface') ?? 'Arial';
      context.font = `${number(properties, 'i') ? 'italic' : 'normal'} ${number(properties, 'b') ? 'bold' : 'normal'} ${fontSize}px ${JSON.stringify(family)}, sans-serif`;
      context.fillStyle = paint(context, properties, width, height) ?? '#000000';
      const value = child(run, 't')?.textContent ?? '';
      const baseline = y - number(properties, 'baseline') / 100000 * fontSize;
      context.fillText(value, x, baseline);
      if (textLayer && value) {
        const span = document.createElement('span');
        span.textContent = value;
        span.style.font = context.font;
        const metrics = context.measureText(value);
        const ascent = metrics.fontBoundingBoxAscent ?? fontSize * .85;
        const lineHeight = ascent + (metrics.fontBoundingBoxDescent ?? fontSize * .35);
        span.style.lineHeight = `${lineHeight}px`;
        span.style.width = `${metrics.width}px`;
        span.style.height = `${lineHeight}px`;
        // Use the same transformed origin as canvas, including rotation and reflection.
        const matrix = context.getTransform().translate(x, baseline - ascent);
        span.style.transform = `matrix(${matrix.a},${matrix.b},${matrix.c},${matrix.d},${matrix.e},${matrix.f})`;
        textLayer.append(span);
      }
      x += context.measureText(value).width;
    }
    y += size * .35;
  }
}

/** Renders the actual generated PPTX, not the original PDF drawing commands. */
async function renderSlide(canvas: HTMLCanvasElement, slide: Slide, signal?: AbortSignal, textLayer?: HTMLDivElement): Promise<void> {
  signal?.throwIfAborted();
  const context = canvas.getContext('2d');
  if (!context) {
    throw new Error('Canvas is unavailable.');
  }
  context.fillStyle = '#ffffff';
  context.fillRect(0, 0, canvas.width, canvas.height);
  context.save();
  context.scale(canvas.width / slide.width, canvas.height / slide.height);
  try {
    const tree = descendants(slide.xml, 'spTree')[0];
    for (const object of Array.from(tree.children)) {
      signal?.throwIfAborted();
      if (object.localName !== 'sp' && object.localName !== 'pic') {
        continue;
      }
      const properties = child(object, 'spPr');
      const frame = child(properties, 'xfrm');
      const width = number(child(frame, 'ext'), 'cx') / emu;
      const height = number(child(frame, 'ext'), 'cy') / emu;
      context.save();
      try {
        context.translate(number(child(frame, 'off'), 'x') / emu + width / 2, number(child(frame, 'off'), 'y') / emu + height / 2);
        context.rotate(number(frame, 'rot') / 60000 * Math.PI / 180);
        context.scale(number(frame, 'flipH') ? -1 : 1, number(frame, 'flipV') ? -1 : 1);
        context.translate(-width / 2, -height / 2);
        if (object.localName === 'pic') {
          const blip = descendants(object, 'blip')[0];
          const id = blip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed') ?? '';
          const path = slide.images.get(id);
          const entry = path ? slide.zip.file(path) : null;
          if (!entry) {
            throw new Error('Missing picture in PowerPoint preview.');
          }
          const bitmap = await createImageBitmap(await entry.async('blob'));
          try {
            context.globalAlpha = number(child(blip, 'alphaModFix'), 'amt', 100000) / 100000;
            context.drawImage(bitmap, 0, 0, width, height);
          } finally {
            bitmap.close();
          }
        } else {
          const geometry = child(properties, 'custGeom');
          const path = new Path2D();
          if (geometry) {
            for (const shapePath of descendants(geometry, 'path')) {
              const sx = width / number(shapePath, 'w', 1), sy = height / number(shapePath, 'h', 1);
              for (const segment of Array.from(shapePath.children)) {
                const points = descendants(segment, 'pt').map(point => [number(point, 'x') * sx, number(point, 'y') * sy]);
                if (segment.localName === 'moveTo') {
                  path.moveTo(points[0][0], points[0][1]);
                } else if (segment.localName === 'lnTo') {
                  path.lineTo(points[0][0], points[0][1]);
                } else if (segment.localName === 'cubicBezTo') {
                  path.bezierCurveTo(points[0][0], points[0][1], points[1][0], points[1][1], points[2][0], points[2][1]);
                } else if (segment.localName === 'close') {
                  path.closePath();
                }
              }
            }
          } else {
            path.rect(0, 0, width, height);
          }
          const fill = paint(context, properties, width, height);
          if (fill) {
            context.fillStyle = fill;
            context.fill(path);
          }
          const line = child(properties, 'ln');
          const stroke = paint(context, line, width, height);
          if (stroke) {
            context.strokeStyle = stroke;
            context.lineWidth = number(line, 'w', emu) / emu;
            context.stroke(path);
          }
          const body = child(object, 'txBody');
          if (body) {
            text(context, body, width, height, textLayer);
          }
        }
      } finally {
        context.restore();
      }
    }
  } finally {
    context.restore();
  }
  signal?.throwIfAborted();
}
