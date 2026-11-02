type Bounds = { x: number; y: number; width: number; height: number; grouped: boolean; rotation: number };
type Run = { id: string; text: string; x: number; y: number; baseline: number; width: number; height: number; fontSize: number; fontFamily: string; color: string; bold: boolean; italic: boolean; underline: boolean; background?: string };
export type DocxInspection = { width: number; height: number; warnings: string[]; media: Record<string, string>; pages: { number: number; id: string; title: string; notes?: string; textObjects: { id: string; name: string; text: string; bounds: Bounds }[]; images: ({ id: string; name: string; partName: string; readOnly?: boolean } & Bounds)[]; runs: Run[]; rules: (Bounds & { fill?: string; strokeWidth: number; color?: string; edges?: ({ width: number; color: string } | null)[] })[] }[] };

export async function createDocxPreview(info: DocxInspection) {
  const pictures = new Map<string, Promise<HTMLImageElement>>();
  function picture(part: string) {
    let promise = pictures.get(part);
    if (!promise) {
      promise = new Promise<HTMLImageElement>((resolve, reject) => {
        const image = new Image();
        image.onload = () => resolve(image);
        image.onerror = () => reject(new Error('Cannot decode document image.'));
        image.src = info.media[part];
      });
      pictures.set(part, promise);
    }
    return promise;
  }
  return {
    width: info.width,
    height: info.height,
    async image(part: string) {
      const uri = info.media[part];
      if (!uri) {
        throw new Error('Image is unavailable.');
      }
      return (await fetch(uri)).blob();
    },
    async render(canvas: HTMLCanvasElement, number: number, signal: AbortSignal, layer?: HTMLElement) {
      signal.throwIfAborted();
      await document.fonts.load('12px "Noto Sans"');
      const page = info.pages[number - 1];
      const context = canvas.getContext('2d')!;
      const scale = canvas.width / info.width;
      context.setTransform(scale, 0, 0, scale, 0, 0);
      context.fillStyle = '#fff';
      context.fillRect(0, 0, info.width, info.height);
      layer?.replaceChildren();
      context.strokeStyle = '#888';
      context.lineWidth = .5;
      for (const rule of page.rules) {
        if (rule.fill && /^[0-9a-f]{6}$/i.test(rule.fill)) {
          context.fillStyle = '#' + rule.fill;
          context.fillRect(rule.x, rule.y, rule.width, rule.height);
        }
        if (rule.edges) {
          const corners = [[rule.x, rule.y], [rule.x + rule.width, rule.y], [rule.x + rule.width, rule.y + rule.height], [rule.x, rule.y + rule.height]];
          rule.edges.forEach((edge, index) => {
            if (!edge || edge.width <= 0) {
              return;
            }
            context.strokeStyle = /^[0-9a-f]{6}$/i.test(edge.color) ? '#' + edge.color : '#000';
            context.lineWidth = edge.width;
            context.lineCap = 'square';
            context.beginPath();
            context.moveTo(...corners[index] as [number, number]);
            context.lineTo(...corners[(index + 1) % 4] as [number, number]);
            context.stroke();
          });
        }
        if (rule.strokeWidth > 0) {
          context.strokeStyle = '#' + (rule.color ?? '888888');
          context.lineWidth = rule.strokeWidth;
          context.strokeRect(rule.x, rule.y, rule.width, rule.height);
        }
      }
      for (const image of page.images) {
        const decoded = await picture(image.partName);
        signal.throwIfAborted();
        context.drawImage(decoded, image.x, image.y, image.width, image.height);
      }
      for (const run of page.runs) {
        const family = run.fontFamily.replace(/["\\]/g, '');
        const font = `${run.italic ? 'italic ' : ''}${run.bold ? 'bold ' : ''}${run.fontSize}px "${family}", "Noto Sans", sans-serif`;
        if (run.background && /^[0-9a-f]{6}$/i.test(run.background)) {
          context.fillStyle = '#' + run.background;
          context.fillRect(run.x, run.y, run.width, run.height);
        }
        context.font = font;
        context.fillStyle = '#' + run.color;
        context.fillText(run.text, run.x, run.baseline);
        if (run.underline) {
          context.fillRect(run.x, run.baseline + 1, run.width, Math.max(.5, run.fontSize / 16));
        }
        if (layer) {
          const span = document.createElement('span');
          span.textContent = run.text;
          span.dataset.textRun = 'true';
          Object.assign(span.style, { position: 'absolute', left: `${run.x * scale}px`, top: `${run.y * scale}px`, font: `${run.italic ? 'italic ' : ''}${run.bold ? 'bold ' : ''}${run.fontSize * scale}px "${family}", "Noto Sans", sans-serif`, lineHeight: `${run.height * scale}px`, whiteSpace: 'pre', color: 'transparent', height: `${run.height * scale}px` });
          layer.append(span);
        }
      }
      if (layer) {
        for (const object of [...page.textObjects.map(item => ({ id: item.id, ...item.bounds })), ...page.images.filter(image => !image.readOnly)]) {
          const box = document.createElement('div');
          box.dataset.objectId = object.id;
          box.className = 'document-object';
          Object.assign(box.style, { position: 'absolute', left: `${object.x * scale}px`, top: `${object.y * scale}px`, width: `${Math.max(3, object.width) * scale}px`, height: `${Math.max(3, object.height) * scale}px` });
          layer.append(box);
        }
      }
    }
  };
}
