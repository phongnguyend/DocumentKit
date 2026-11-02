import { expandTables } from './pptxTables';
import type JSZip from 'jszip';

const emu = 12700;
const child = (node: Element | null, name: string) => Array.from(node?.children ?? []).find(item => item.localName === name) ?? null;
const descendants = (node: Element | Document, name: string) => Array.from(node.getElementsByTagNameNS('*', name));
const number = (node: Element | null, name: string, fallback = 0) => {
  const raw = node?.getAttribute(name);
  const value = raw === 'true' ? 1 : raw === 'false' ? 0 : Number(raw ?? fallback);
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

/** Approximate DrawingML preview; unsupported content remains in the saved package. */
export async function createPptxPreview(bytes: ArrayBuffer) {
  const { default: Zip } = await import('jszip');
  const zip = await Zip.loadAsync(bytes);
  const rootRelations = await xml(zip, '_rels/.rels');
  const root = descendants(rootRelations, 'Relationship').find(rel => rel.getAttribute('Type')?.endsWith('/officeDocument'));
  if (!root || root.getAttribute('TargetMode') === 'External') {
    throw new Error('Missing internal presentation relationship.');
  }
  const resolve = (part: string, target: string) => new URL(target, 'https://package.invalid/' + part).pathname.slice(1);
  const relPath = (part: string) => part.replace(/([^/]+)$/, '_rels/$1.rels');
  const main = resolve('', root.getAttribute('Target') ?? '');
  const presentation = await xml(zip, main);
  const mainRelations = await xml(zip, relPath(main));
  const slidePaths = descendants(presentation, 'sldId').map(id => {
    const relationship = descendants(mainRelations, 'Relationship').find(rel => rel.getAttribute('Id') === id.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id'));
    if (!relationship || relationship.getAttribute('TargetMode') === 'External') {
      throw new Error('Invalid slide relationship.');
    }
    return resolve(main, relationship.getAttribute('Target') ?? '');
  });
  const size = descendants(presentation, 'sldSz')[0];
  const width = number(size, 'cx') / emu;
  const height = number(size, 'cy') / emu;
  const count = descendants(presentation, 'sldId').length;
  if (width <= 0 || height <= 0 || count === 0) {
    throw new Error('Invalid PowerPoint slide size or slide count.');
  }
  return {
    width, height, count,
    async image(part: string) {
      const entry = zip.file(part);
      if (!entry) {
        throw new Error('Image is missing from the presentation.');
      }
      return entry.async('blob');
    },
    async render(canvas: HTMLCanvasElement, page: number, signal?: AbortSignal, textLayer?: HTMLDivElement) {
      if (page < 1 || page > count) {
        throw new Error('Invalid PowerPoint preview slide.');
      }
      const part = slidePaths[page - 1];
      const relationships = zip.file(relPath(part)) ? await xml(zip, relPath(part)) : new DOMParser().parseFromString('<Relationships/>', 'application/xml');
      const images = new Map<string, string>();
      for (const relationship of descendants(relationships, 'Relationship')) {
        if (relationship.getAttribute('Type')?.endsWith('/image') && relationship.getAttribute('TargetMode') !== 'External') {
          const target = new URL(relationship.getAttribute('Target') ?? '', 'https://pptx.invalid/' + part);
          images.set(relationship.getAttribute('Id') ?? '', target.pathname.slice(1));
        }
      }
      const slide = { zip, xml: await xml(zip, part), width, height, images, notes: [] };
      for (const picture of descendants(slide.xml, 'pic')) {
        const id = child(child(picture, 'nvPicPr'), 'cNvPr')?.getAttribute('id');
        if (id) {
          picture.setAttribute('data-object-id', id);
        }
      }
      for (const shape of descendants(slide.xml, 'sp')) {
        const id = child(child(shape, 'nvSpPr'), 'cNvPr')?.getAttribute('id');
        if (id) {
          child(shape, 'txBody')?.setAttribute('data-object-id', id);
        }
      }
      const loadRelated = async (source: string, kind: string) => {
        if (!zip.file(relPath(source))) {
          return null;
        }
        const rels = await xml(zip, relPath(source));
        const rel = descendants(rels, 'Relationship').find(item => item.getAttribute('Type')?.endsWith('/' + kind) && item.getAttribute('TargetMode') !== 'External');
        if (!rel) {
          return null;
        }
        const path = resolve(source, rel.getAttribute('Target') ?? '');
        return { path, doc: await xml(zip, path), rels };
      };
      const layout = await loadRelated(part, 'slideLayout');
      const master = layout ? await loadRelated(layout.path, 'slideMaster') : null;
      const theme = master ? await loadRelated(master.path, 'theme') : null;
      // Office stores a resolved drawing for SmartArt alongside the editable data model.
      // Use its authored geometry rather than attempting to recompute the layout algorithm.
      for (const graphic of descendants(slide.xml, 'graphicFrame')) {
        const ids = descendants(graphic, 'relIds')[0];
        if (!ids) {
          continue;
        }
        const ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
        const dataRel = descendants(relationships, 'Relationship').find(item => item.getAttribute('Id') === ids.getAttributeNS(ns, 'dm') && item.getAttribute('TargetMode') !== 'External');
        if (!dataRel) {
          continue;
        }
        const data = await xml(zip, resolve(part, dataRel.getAttribute('Target') ?? ''));
        const drawingId = descendants(data, 'dataModelExt')[0]?.getAttribute('relId');
        const drawingRel = descendants(relationships, 'Relationship').find(item => item.getAttribute('Id') === drawingId && item.getAttribute('Type')?.endsWith('/diagramDrawing') && item.getAttribute('TargetMode') !== 'External');
        if (!drawingRel) {
          continue;
        }
        const drawing = await xml(zip, resolve(part, drawingRel.getAttribute('Target') ?? ''));
        const drawingTree = descendants(drawing, 'spTree')[0];
        const frame = child(graphic, 'xfrm');
        if (!drawingTree || !frame) {
          continue;
        }
        const group = slide.xml.createElementNS(graphic.namespaceURI, 'p:grpSp');
        const props = slide.xml.createElementNS(graphic.namespaceURI, 'p:grpSpPr');
        const transform = slide.xml.createElementNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'a:xfrm');
        for (const item of Array.from(frame.children)) {
          transform.append(item.cloneNode(true));
        }
        const offset = slide.xml.createElementNS(transform.namespaceURI, 'a:chOff');
        offset.setAttribute('x', '0');
        offset.setAttribute('y', '0');
        const ext = slide.xml.createElementNS(transform.namespaceURI, 'a:chExt');
        ext.setAttribute('cx', child(frame, 'ext')?.getAttribute('cx') ?? '1');
        ext.setAttribute('cy', child(frame, 'ext')?.getAttribute('cy') ?? '1');
        transform.append(offset, ext);
        props.append(transform);
        group.append(props, ...Array.from(drawingTree.children).filter(item => !['nvGrpSpPr', 'grpSpPr'].includes(item.localName)).map(item => item.cloneNode(true)));
        graphic.replaceWith(group);
      }
      const tree = descendants(slide.xml, 'spTree')[0];
      const decorations: Element[] = [];
      for (const layer of [master, layout]) {
        if (!layer || slide.xml.documentElement.getAttribute('showMasterSp') === '0') {
          continue;
        }
        const layerRels = zip.file(relPath(layer.path)) ? await xml(zip, relPath(layer.path)) : null;
        for (const object of Array.from(descendants(layer.doc, 'spTree')[0]?.children ?? [])) {
          if (descendants(object, 'ph').length || !['sp', 'pic', 'grpSp', 'cxnSp'].includes(object.localName)) {
            continue;
          }
          const clone = object.cloneNode(true) as Element;
          for (const blip of [...descendants(clone, 'blip'), ...descendants(clone, 'svgBlip')]) {
            const ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
            const id = blip.getAttributeNS(ns, 'embed');
            const rel = layerRels ? descendants(layerRels, 'Relationship').find(item => item.getAttribute('Id') === id && item.getAttribute('TargetMode') !== 'External') : null;
            if (rel) {
              const unique = layer.path + ':' + id;
              images.set(unique, resolve(layer.path, rel.getAttribute('Target') ?? ''));
              blip.setAttributeNS(ns, 'r:embed', unique);
            }
          }
          decorations.push(clone);
        }
      }
      tree.prepend(...decorations);
      if (!descendants(slide.xml, 'bg').length) {
        const background = (layout ? descendants(layout.doc, 'bg')[0] : null) ?? (master ? descendants(master.doc, 'bg')[0] : null);
        if (background) {
          descendants(slide.xml, 'cSld')[0].prepend(background.cloneNode(true));
        }
      }
      const tableStyles = await loadRelated(main, 'tableStyles');
      expandTables(slide.xml, tableStyles?.doc ?? null);
      const placeholder = (shape: Element, doc?: Document) => {
        const ph = descendants(shape, 'ph')[0];
        if (!ph || !doc) {
          return null;
        }
        const shapes = descendants(doc, 'sp');
        return shapes.find(candidate => {
          const other = descendants(candidate, 'ph')[0];
          return other && (other.getAttribute('idx') ?? '0') === (ph.getAttribute('idx') ?? '0');
        }) ?? shapes.find(candidate => descendants(candidate, 'ph')[0]?.getAttribute('type') === ph.getAttribute('type')) ?? null;
      };
      const merge = (target: Element, source: Element | null) => {
        if (!source) {
          return;
        }
        for (const attribute of Array.from(source.attributes)) {
          if (!target.hasAttribute(attribute.name)) {
            target.setAttribute(attribute.name, attribute.value);
          }
        }
        for (const item of Array.from(source.children)) {
          const choices = [
            ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'],
            ['buNone', 'buChar', 'buAutoNum', 'buBlip'],
            ['buClr', 'buClrTx'], ['buSzPct', 'buSzPts', 'buSzTx'], ['buFont', 'buFontTx'],
          ];
          const exclusive = choices.find(group => group.includes(item.localName));
          if (!child(target, item.localName) && !exclusive?.some(name => child(target, name))) {
            target.append(item.cloneNode(true));
          }
        }
      };
      const styleTemplate = (reference: Element | null, list: string) => {
        const index = number(reference, 'idx');
        if (!theme || index <= 0) {
          return null;
        }
        const source = Array.from(descendants(theme.doc, index >= 1001 ? 'bgFillStyleLst' : list)[0]?.children ?? [])[index >= 1001 ? index - 1001 : index - 1];
        if (!source) {
          return null;
        }
        const copy = source.cloneNode(true) as Element;
        for (const color of descendants(copy, 'schemeClr').filter(item => item.getAttribute('val') === 'phClr')) {
          const replacement = reference?.firstElementChild?.cloneNode(true) as Element | undefined;
          if (replacement) {
            replacement.append(...Array.from(color.children).map(item => item.cloneNode(true)));
            color.replaceWith(replacement);
          }
        }
        return copy;
      };
      for (const shape of [...descendants(slide.xml, 'sp'), ...descendants(slide.xml, 'cxnSp')]) {
        const properties = child(shape, 'spPr');
        const style = child(shape, 'style');
        if (!properties || !style) {
          continue;
        }
        const fill = styleTemplate(child(style, 'fillRef'), 'fillStyleLst');
        if (fill && !Array.from(properties.children).some(item => ['noFill', 'solidFill', 'gradFill', 'blipFill', 'grpFill', 'pattFill'].includes(item.localName))) {
          properties.append(fill);
        }
        const line = styleTemplate(child(style, 'lnRef'), 'lnStyleLst');
        if (line) {
          const own = child(properties, 'ln');
          if (own) {
            merge(own, line);
          } else {
            properties.append(line);
          }
        }
      }
      for (const shape of descendants(slide.xml, 'sp')) {
        const inherited = placeholder(shape, layout?.doc);
        const base = placeholder(inherited ?? shape, master?.doc);
        const props = child(shape, 'spPr');
        if (props) {
          merge(props, child(inherited, 'spPr'));
          merge(props, child(base, 'spPr'));
        }
        const body = child(shape, 'txBody');
        if (!body) {
          continue;
        }
        const bodyPr = child(body, 'bodyPr');
        if (bodyPr) {
          merge(bodyPr, child(child(inherited, 'txBody'), 'bodyPr'));
          merge(bodyPr, child(child(base, 'txBody'), 'bodyPr'));
        }
        const type = descendants(shape, 'ph')[0]?.getAttribute('type') ?? 'other';
        const style = type === 'title' || type === 'ctrTitle' ? 'titleStyle' : type === 'body' || type === 'subTitle' || type === 'obj' ? 'bodyStyle' : 'otherStyle';
        for (const paragraph of Array.from(body.children).filter(item => item.localName === 'p')) {
          let pPr = child(paragraph, 'pPr');
          if (!pPr) {
            pPr = slide.xml.createElementNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'a:pPr');
            paragraph.prepend(pPr);
          }
          const level = `lvl${number(pPr, 'lvl') + 1}pPr`;
          const defaults = [child(child(body, 'lstStyle'), level), child(child(child(inherited, 'txBody'), 'lstStyle'), level), child(child(child(base, 'txBody'), 'lstStyle'), level), child(master ? descendants(master.doc, style)[0] : null, level), child(descendants(presentation, 'defaultTextStyle')[0], level)];
          for (const fallback of defaults) {
            merge(pPr, fallback);
          }
          for (const run of Array.from(paragraph.children).filter(item => item.localName === 'r' || item.localName === 'fld')) {
            let rPr = child(run, 'rPr');
            if (!rPr) {
              rPr = slide.xml.createElementNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'a:rPr');
              run.prepend(rPr);
            }
            if (child(rPr, 'hlinkClick')) {
              if (!child(rPr, 'solidFill') && !child(rPr, 'gradFill') && !child(rPr, 'noFill')) {
                const fill = slide.xml.createElementNS(rPr.namespaceURI, 'a:solidFill');
                const linkColor = slide.xml.createElementNS(rPr.namespaceURI, 'a:schemeClr');
                linkColor.setAttribute('val', 'hlink');
                fill.append(linkColor);
                rPr.append(fill);
              }
              if (!rPr.hasAttribute('u')) {
                rPr.setAttribute('u', 'sng');
              }
            }
            const fontRef = child(child(shape, 'style'), 'fontRef');
            const fontDefaults = slide.xml.createElementNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'a:rPr');
            if (fontRef?.firstElementChild) {
              const fill = slide.xml.createElementNS(fontDefaults.namespaceURI, 'a:solidFill');
              fill.append(fontRef.firstElementChild.cloneNode(true));
              fontDefaults.append(fill);
            }
            merge(rPr, fontDefaults);
            merge(rPr, child(pPr, 'defRPr'));
            for (const fallback of defaults) {
              merge(rPr, child(fallback, 'defRPr'));
            }
          }
        }
      }
      for (const fill of descendants(slide.xml, 'grpFill')) {
        let group = fill.parentElement?.parentElement?.parentElement;
        while (group) {
          if (group.localName === 'grpSp') {
            const inherited = Array.from(child(group, 'grpSpPr')?.children ?? []).find(item => ['solidFill', 'gradFill', 'noFill'].includes(item.localName));
            if (inherited) {
              fill.replaceWith(inherited.cloneNode(true));
              break;
            }
          }
          group = group.parentElement;
        }
      }
      // Theme colors and fonts must be resolved before painting inherited shapes.
      const map = master ? descendants(master.doc, 'clrMap')[0] : null;
      const override = layout ? descendants(layout.doc, 'overrideClrMapping')[0] : null;
      const slideMap = descendants(slide.xml, 'overrideClrMapping')[0];
      const scheme = theme ? descendants(theme.doc, 'clrScheme')[0] : null;
      for (const node of descendants(slide.xml, 'schemeClr')) {
        const key = node.getAttribute('val') ?? '';
        const mapped = slideMap?.getAttribute(key) ?? override?.getAttribute(key) ?? map?.getAttribute(key) ?? ({ tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' } as Record<string, string>)[key] ?? key;
        const source = child(scheme, mapped)?.firstElementChild;
        const replacement = slide.xml.createElementNS(node.namespaceURI, 'a:srgbClr');
        replacement.setAttribute('val', source?.getAttribute('lastClr') ?? source?.getAttribute('val') ?? (mapped.startsWith('lt') ? 'FFFFFF' : '000000'));
        replacement.append(...Array.from(node.childNodes).map(item => item.cloneNode(true)));
        node.replaceWith(replacement);
      }
      for (const node of descendants(slide.xml, 'latin')) {
        const face = node.getAttribute('typeface') ?? '';
        if (face.startsWith('+')) {
          const font = theme ? descendants(theme.doc, face.startsWith('+mj') ? 'majorFont' : 'minorFont')[0] : null;
          node.setAttribute('typeface', child(font, 'latin')?.getAttribute('typeface') ?? 'Arial');
        }
      }
      await renderSlide(canvas, slide, signal, textLayer);
    },
  };
}

function color(node: Element | null): string {
  const rgb = node?.getAttribute('val') ?? '000000';
  if (!/^[0-9a-f]{6}$/i.test(rgb)) {
    return '#000000';
  }
  let channels = [0, 2, 4].map(offset => parseInt(rgb.slice(offset, offset + 2), 16));
  for (const transform of Array.from(node?.children ?? [])) {
    const value = number(transform, 'val') / 100000;
    if (transform.localName === 'tint') {
      channels = channels.map(channel => channel + (255 - channel) * value);
    } else if (transform.localName === 'shade' || transform.localName === 'lumMod') {
      channels = channels.map(channel => channel * value);
    } else if (transform.localName === 'lumOff') {
      channels = channels.map(channel => channel + 255 * value);
    }
  }
  return `rgba(${channels.map(channel => Math.round(Math.max(0, Math.min(255, channel)))).join(',')},${number(child(node, 'alpha'), 'val', 100000) / 100000})`;
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
  const settings = child(body, 'bodyPr');
  const left = number(settings, 'lIns', 91440) / emu;
  const right = number(settings, 'rIns', 91440) / emu;
  const top = number(settings, 'tIns', 45720) / emu;
  const bottom = number(settings, 'bIns', 45720) / emu;
  const scale = number(child(settings, 'normAutofit'), 'fontScale', 100000) / 100000;
  type Fragment = { value: string; font: string; size: number; width: number; props: Element | null; markerX?: number; markerColor?: string };
  type Line = { fragments: Fragment[]; width: number; size: number; align: string; x: number; advance: number; before?: number; marker?: Fragment };
  const lines: Line[] = [];
  const counters = new Map<number, number>();
  for (const paragraph of Array.from(body.children).filter(item => item.localName === 'p')) {
    const pPr = child(paragraph, 'pPr');
    const paragraphStart = lines.length;
    const margin = number(pPr, 'marL') / emu;
    const available = Math.max(1, width - left - right - margin - number(pPr, 'marR') / emu);
    const align = pPr?.getAttribute('algn') ?? 'l';
    let line: Line = { fragments: [], width: 0, size: 12 * scale, align, x: left + margin, advance: 0 };
    const flush = () => {
      const spacing = child(pPr, 'lnSpc');
      line.advance = child(spacing, 'spcPts') ? number(child(spacing, 'spcPts'), 'val') / 100 : line.size * 1.2 * Math.max(0, number(child(spacing, 'spcPct'), 'val', 100000) - number(child(settings, 'normAutofit'), 'lnSpcReduction')) / 100000;
      if (align === 'ctr') {
        line.x += (available - line.width) / 2;
      } else if (align === 'r') {
        line.x += available - line.width;
      }
      lines.push(line);
      line = { fragments: [], width: 0, size: 12 * scale, align, x: left + margin, advance: 0 };
    };
    let pending: Fragment[] = [];
    const commit = () => {
      const total = pending.reduce((sum, fragment) => sum + fragment.width, 0);
      if (settings?.getAttribute('wrap') !== 'none' && total <= available && line.width + total > available && line.fragments.length && pending.some(fragment => fragment.value.trim())) {
        flush();
      }
      for (const fragment of pending) {
        const { value: token, font, size, props } = fragment;
        context.font = font;
        const tokenWidth = fragment.width;
        if (settings?.getAttribute('wrap') !== 'none' && tokenWidth > available && token.trim()) {
          // Long URLs and unspaced text must still fit the text frame. Keep grapheme
          // clusters intact so surrogate pairs and combining marks are not separated.
          const characters = Array.from(new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(token), item => item.segment);
          let offset = 0;
          while (offset < characters.length) {
            const remaining = available - line.width;
            let low = 1, high = characters.length - offset, count = 0;
            while (low <= high) {
              const middle = Math.floor((low + high) / 2);
              if (context.measureText(characters.slice(offset, offset + middle).join('')).width <= remaining) {
                count = middle;
                low = middle + 1;
              } else {
                high = middle - 1;
              }
            }
            if (!count && line.fragments.length) {
              flush();
              continue;
            }
            count = Math.max(1, count);
            const piece = characters.slice(offset, offset + count).join('');
            const pieceWidth = context.measureText(piece).width;
            line.fragments.push({ value: piece, font, size, width: pieceWidth, props });
            line.width += pieceWidth;
            line.size = Math.max(line.size, size);
            offset += count;
            if (offset < characters.length) {
              flush();
            }
          }
          continue;
        }
        if (settings?.getAttribute('wrap') !== 'none' && line.width + tokenWidth > available && line.fragments.length && token.trim()) {
          flush();
        }
        line.fragments.push({ value: token, font, size, width: tokenWidth, props });
        line.width += tokenWidth;
        line.size = Math.max(line.size, size);
      }
      pending = [];
    };
    for (const run of Array.from(paragraph.children)) {
      if (run.localName === 'br') {
        commit();
        flush();
        continue;
      }
      if (run.localName !== 'r' && run.localName !== 'fld') {
        continue;
      }
      const props = child(run, 'rPr');
      const size = number(props, 'sz', 1800) / 100 * scale;
      const family = child(props, 'latin')?.getAttribute('typeface') ?? 'Arial';
      const font = `${number(props, 'i') ? 'italic' : 'normal'} ${number(props, 'b') ? 'bold' : 'normal'} ${size}px ${JSON.stringify(family)}, sans-serif`;
      context.font = font;
      let value = child(run, 't')?.textContent ?? '';
      if (child(props, 'sym')?.getAttribute('typeface') === 'Wingdings') {
        // Map only known symbol codepoints; Latin characters in the same run stay Latin.
        const symbols: Record<string, string> = { '\uf0e8': '\u2794', '\uf0a7': '\u25aa', '\uf0fc': '\u2713', '\uf0fb': '\u2717', '\uf04c': '\u2639\ufe0e' };
        value = Array.from(value, character => symbols[character] ?? character).join('');
      }
      const tokens = settings?.getAttribute('wrap') === 'none' ? [value] : value.match(/[^ \t\r\n]+|[ \t\r\n]+/gu) ?? [];
      for (const token of tokens) {
        context.font = font;
        const fragment = { value: token, font, size, width: context.measureText(token).width, props };
        if (/^[ \t\r\n]+$/u.test(token)) {
          commit();
          pending.push(fragment);
          commit();
        } else {
          pending.push(fragment);
        }
      }
    }
    commit();
    flush();
    const first = lines[paragraphStart];
    const spacing = (name: string) => {
      const node = child(pPr, name);
      return child(node, 'spcPts') ? number(child(node, 'spcPts'), 'val') / 100 : first.size * number(child(node, 'spcPct'), 'val') / 100000;
    };
    first.before = spacing('spcBef');
    lines[lines.length - 1].advance += spacing('spcAft');
    const firstRun = first.fragments[0];
    const auto = child(pPr, 'buAutoNum');
    const bullet = child(pPr, 'buChar');
    const level = number(pPr, 'lvl');
    for (const key of counters.keys()) {
      if (key > level) {
        counters.delete(key);
      }
    }
    if (firstRun && !child(pPr, 'buNone') && (auto || bullet)) {
      let value = bullet?.getAttribute('char') ?? '';
      if (auto) {
        const count = auto.hasAttribute('startAt') ? number(auto, 'startAt', 1) : (counters.get(level) ?? 0) + 1;
        counters.set(level, count);
        const type = auto.getAttribute('type') ?? 'arabicPeriod';
        value = type.startsWith('alpha') ? String.fromCharCode((type.startsWith('alphaUc') ? 65 : 97) + (count - 1) % 26) : String(count);
        value = type.endsWith('ParenBoth') ? `(${value})` : type.endsWith('ParenR') ? `${value})` : type.endsWith('Plain') ? value : `${value}.`;
      }
      // Common Wingdings bullets need Unicode equivalents on systems without that font.
      const face = child(pPr, 'buFont')?.getAttribute('typeface') ?? '';
      if (/Wingdings/i.test(face) && ['\u00a7', '\ufffd', '\uf0a7'].includes(value)) {
        value = '\u25aa';
      }
      const size = child(pPr, 'buSzPts') ? number(child(pPr, 'buSzPts'), 'val') / 100 : firstRun.size * number(child(pPr, 'buSzPct'), 'val', 100000) / 100000;
      first.marker = { ...firstRun, value, size, font: `${size}px Arial, sans-serif`, markerX: left + margin + number(pPr, 'indent') / emu, markerColor: child(pPr, 'buClr') ? color(child(child(pPr, 'buClr'), 'srgbClr')) : undefined };
    }
  }
  const objectBounds = { left: Infinity, top: Infinity, right: -Infinity, bottom: -Infinity };
  const total = lines.reduce((sum, line) => sum + line.advance + (line.before ?? 0), 0);
  const anchor = settings?.getAttribute('anchor');
  let y = top + (anchor === 'ctr' ? (height - top - bottom - total) / 2 : anchor === 'b' ? height - top - bottom - total : 0);
  for (const line of lines) {
    y += line.before ?? 0;
    let x = line.x;
    for (const fragment of (line.marker ? [line.marker, ...line.fragments] : line.fragments)) {
      const drawX = fragment.markerX ?? x;
      context.font = fragment.font;
      context.fillStyle = fragment.markerColor ?? paint(context, fragment.props, width, height) ?? '#000000';
      const baseline = y + line.size * .9 - number(fragment.props, 'baseline') / 100000 * fragment.size;
      context.fillText(fragment.value, drawX, baseline);
      const underline = fragment.props?.getAttribute('u');
      if (fragment.markerX === undefined && underline && underline !== 'none') {
        context.fillRect(drawX, baseline + fragment.size * .08, fragment.width, Math.max(.5, fragment.size / 18));
      }
      if (textLayer && fragment.value) {
        const span = document.createElement('span');
        span.textContent = fragment.value;
        const objectId = body.getAttribute('data-object-id');
        if (objectId) {
          span.dataset.textObjectId = objectId;
        }
        span.style.font = context.font;
        const ascent = context.measureText(fragment.value).fontBoundingBoxAscent ?? fragment.size * .9;
        span.style.lineHeight = `${fragment.size * 1.2}px`;
        span.style.width = `${fragment.width}px`;
        span.style.height = `${fragment.size * 1.2}px`;
        objectBounds.left = Math.min(objectBounds.left, drawX);
        objectBounds.top = Math.min(objectBounds.top, baseline - ascent);
        objectBounds.right = Math.max(objectBounds.right, drawX + fragment.width);
        objectBounds.bottom = Math.max(objectBounds.bottom, baseline - ascent + fragment.size * 1.2);
        const matrix = context.getTransform().translate(drawX, baseline - ascent);
        span.style.transform = `matrix(${matrix.a},${matrix.b},${matrix.c},${matrix.d},${matrix.e},${matrix.f})`;
        textLayer.append(span);
      }
      if (fragment.markerX === undefined) {
        x += fragment.width;
      }
    }
    y += line.advance;
  }
  const objectId = body.getAttribute('data-object-id');
  if (textLayer && objectId && Number.isFinite(objectBounds.left)) {
    const hit = document.createElement('span');
    hit.dataset.objectId = objectId;
    hit.dataset.kind = 'text';
    hit.className = 'text-object-hit';
    hit.style.width = `${objectBounds.right - objectBounds.left}px`;
    hit.style.height = `${objectBounds.bottom - objectBounds.top}px`;
    const matrix = context.getTransform().translate(objectBounds.left, objectBounds.top);
    hit.style.transform = `matrix(${matrix.a},${matrix.b},${matrix.c},${matrix.d},${matrix.e},${matrix.f})`;
    textLayer.append(hit);
  }
}

/** SVG extensions often omit raster fallbacks and explicit pixel dimensions. */
async function decodePicture(blob: Blob, path: string): Promise<ImageBitmap> {
  if (!path.toLowerCase().endsWith('.svg')) {
    return createImageBitmap(blob);
  }
  const xml = new DOMParser().parseFromString(await blob.text(), 'image/svg+xml');
  const root = xml.documentElement;
  if (root.localName !== 'svg' || xml.querySelector('parsererror')) {
    throw new Error('Invalid SVG picture.');
  }
  if (!root.hasAttribute('width') || !root.hasAttribute('height')) {
    const viewBox = (root.getAttribute('viewBox') ?? '').trim().split(/[\s,]+/).map(Number);
    const width = viewBox.length === 4 && viewBox[2] > 0 ? viewBox[2] : 300;
    const height = viewBox.length === 4 && viewBox[3] > 0 ? viewBox[3] : 150;
    const scale = Math.min(1, 4096 / Math.max(width, height));
    root.setAttribute('width', String(Math.max(1, width * scale)));
    root.setAttribute('height', String(Math.max(1, height * scale)));
  }
  const url = URL.createObjectURL(new Blob([new XMLSerializer().serializeToString(xml)], { type: 'image/svg+xml' }));
  try {
    const image = new Image();
    image.src = url;
    await image.decode();
    return await createImageBitmap(image);
  } finally {
    URL.revokeObjectURL(url);
  }
}

/** Paints supported slide DrawingML in package order. */
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
    const background = descendants(slide.xml, 'bgPr')[0];
    const backgroundFill = paint(context, background ?? null, slide.width, slide.height);
    if (backgroundFill) {
      context.fillStyle = backgroundFill;
      context.fillRect(0, 0, slide.width, slide.height);
    }
    const tree = descendants(slide.xml, 'spTree')[0];
    const renderObjects = async (tree: Element) => {
      for (const object of Array.from(tree.children)) {
        signal?.throwIfAborted();
        if (object.localName === 'grpSp') {
          const frame = child(child(object, 'grpSpPr'), 'xfrm');
          const w = number(child(frame, 'ext'), 'cx') / emu;
          const h = number(child(frame, 'ext'), 'cy') / emu;
          const cw = number(child(frame, 'chExt'), 'cx') / emu;
          const ch = number(child(frame, 'chExt'), 'cy') / emu;
          context.save();
          try {
            context.translate(number(child(frame, 'off'), 'x') / emu + w / 2, number(child(frame, 'off'), 'y') / emu + h / 2);
            context.rotate(number(frame, 'rot') / 60000 * Math.PI / 180);
            context.scale(number(frame, 'flipH') ? -1 : 1, number(frame, 'flipV') ? -1 : 1);
            context.translate(-w / 2, -h / 2);
            context.scale(cw ? w / cw : 1, ch ? h / ch : 1);
            context.translate(-number(child(frame, 'chOff'), 'x') / emu, -number(child(frame, 'chOff'), 'y') / emu);
            await renderObjects(object);
          } finally {
            context.restore();
          }
          continue;
        }
        if (object.localName !== 'sp' && object.localName !== 'pic' && object.localName !== 'cxnSp') {
          continue;
        }
        const properties = child(object, 'spPr');
        const frame = child(properties, 'xfrm');
        if (!frame) {
          continue;
        }
        const width = number(child(frame, 'ext'), 'cx') / emu;
        const height = number(child(frame, 'ext'), 'cy') / emu;
        context.save();
        try {
          context.translate(number(child(frame, 'off'), 'x') / emu + width / 2, number(child(frame, 'off'), 'y') / emu + height / 2);
          context.rotate(number(frame, 'rot') / 60000 * Math.PI / 180);
          context.scale(number(frame, 'flipH') ? -1 : 1, number(frame, 'flipV') ? -1 : 1);
          context.translate(-width / 2, -height / 2);
          if (object.localName === 'pic') {
            const objectId = object.getAttribute('data-object-id');
            if (textLayer && objectId) {
              const hit = document.createElement('span');
              hit.dataset.objectId = objectId;
              hit.dataset.kind = 'image';
              hit.style.width = `${width}px`;
              hit.style.height = `${height}px`;
              const matrix = context.getTransform();
              hit.style.transform = `matrix(${matrix.a},${matrix.b},${matrix.c},${matrix.d},${matrix.e},${matrix.f})`;
              textLayer.append(hit);
            }
            const blip = descendants(object, 'blip')[0];
            const ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
            const references = [descendants(object, 'svgBlip')[0]?.getAttributeNS(ns, 'embed'), blip?.getAttributeNS(ns, 'embed')];
            const candidates = references.map(id => id ? slide.images.get(id) : undefined).filter((path): path is string => !!path && !!slide.zip.file(path));
            const entry = candidates.length > 0;
            if (!entry) {
              throw new Error('Missing picture in PowerPoint preview.');
            }
            let bitmap: ImageBitmap | undefined;
            for (const path of candidates) {
              try {
                bitmap = await decodePicture(await slide.zip.file(path)!.async('blob'), path);
                break;
              } catch {
                // Try a raster fallback when the SVG extension cannot be decoded.
              }
            }
            if (!bitmap) {
              continue;
            }
            try {
              context.globalAlpha = number(child(blip, 'alphaModFix'), 'amt', 100000) / 100000;
              const crop = child(child(object, 'blipFill'), 'srcRect');
              const l = number(crop, 'l') / 100000, t = number(crop, 't') / 100000;
              const r = number(crop, 'r') / 100000, b = number(crop, 'b') / 100000;
              context.drawImage(bitmap, l * bitmap.width, t * bitmap.height, bitmap.width * (1 - l - r), bitmap.height * (1 - t - b), 0, 0, width, height);
            } finally {
              bitmap.close();
            }
          } else {
            const geometry = child(properties, 'custGeom');
            const path = new Path2D();
            let connector: number[][] | null = null;
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
              if (object.localName === 'cxnSp' || child(properties, 'prstGeom')?.getAttribute('prst') === 'line') {
                const preset = child(properties, 'prstGeom');
                const adjustment = preset ? descendants(preset, 'gd').find(item => item.getAttribute('name') === 'adj1')?.getAttribute('fmla') : null;
                const middle = width * Number(adjustment?.replace('val ', '') ?? 50000) / 100000;
                connector = preset?.getAttribute('prst') === 'bentConnector2' ? [[0, 0], [width, 0], [width, height]]
                  : preset?.getAttribute('prst') === 'bentConnector3' ? [[0, 0], [middle, 0], [middle, height], [width, height]] : [[0, 0], [width, height]];
                path.moveTo(0, 0);
                for (const point of connector.slice(1)) {
                  path.lineTo(point[0], point[1]);
                }
              } else if (child(properties, 'prstGeom')?.getAttribute('prst') === 'round2DiagRect') {
                const adjustment = descendants(child(properties, 'prstGeom')!, 'gd').find(item => item.getAttribute('name') === 'adj1')?.getAttribute('fmla');
                const radius = Math.min(width, height) * Math.min(.5, Math.max(0, Number(adjustment?.replace('val ', '') ?? 16667) / 100000));
                path.moveTo(radius, 0);
                path.lineTo(width, 0);
                path.lineTo(width, height - radius);
                path.quadraticCurveTo(width, height, width - radius, height);
                path.lineTo(0, height);
                path.lineTo(0, radius);
                path.quadraticCurveTo(0, 0, radius, 0);
                path.closePath();
              } else if (child(properties, 'prstGeom')?.getAttribute('prst') === 'leftBrace') {
                path.moveTo(width, 0);
                path.bezierCurveTo(0, 0, width, height / 2, 0, height / 2);
                path.bezierCurveTo(width, height / 2, 0, height, width, height);
              } else if (child(properties, 'prstGeom')?.getAttribute('prst') === 'ellipse') {
                path.ellipse(width / 2, height / 2, width / 2, height / 2, 0, 0, Math.PI * 2);
              } else {
                path.rect(0, 0, width, height);
              }
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
              if (connector) {
                for (const end of ['headEnd', 'tailEnd']) {
                  const marker = child(line, end);
                  if (marker?.getAttribute('type') !== 'triangle' && marker?.getAttribute('type') !== 'arrow') {
                    continue;
                  }
                  const points = end === 'headEnd' ? connector : [...connector].reverse();
                  const tip = points[0];
                  const next = points.find(point => Math.hypot(point[0] - tip[0], point[1] - tip[1]) > .01);
                  if (!next) {
                    continue;
                  }
                  context.save();
                  context.translate(tip[0], tip[1]);
                  context.rotate(Math.atan2(tip[1] - next[1], tip[0] - next[0]));
                  const length = context.lineWidth * (marker.getAttribute('len') === 'lg' ? 5 : marker.getAttribute('len') === 'sm' ? 2 : 3);
                  const half = context.lineWidth * (marker.getAttribute('w') === 'lg' ? 2.5 : marker.getAttribute('w') === 'sm' ? 1 : 1.5);
                  context.beginPath();
                  context.moveTo(-length, -half);
                  context.lineTo(0, 0);
                  context.lineTo(-length, half);
                  if (marker.getAttribute('type') === 'triangle') {
                    context.closePath();
                    context.fillStyle = stroke;
                    context.fill();
                  } else {
                    context.stroke();
                  }
                  context.restore();
                }
              }
            }
            const body = child(object, 'txBody');
            if (body) {
              context.save();
              try {
                // DrawingML flips geometry; text remains readable inside the shape.
                context.translate(width / 2, height / 2);
                context.scale(number(frame, 'flipH') ? -1 : 1, number(frame, 'flipV') ? -1 : 1);
                context.translate(-width / 2, -height / 2);
                text(context, body, width, height, textLayer);
              } finally {
                context.restore();
              }
            }
          }
        } finally {
          context.restore();
        }
      }
    };
    await renderObjects(tree);
  } finally {
    context.restore();
  }
  signal?.throwIfAborted();
}
