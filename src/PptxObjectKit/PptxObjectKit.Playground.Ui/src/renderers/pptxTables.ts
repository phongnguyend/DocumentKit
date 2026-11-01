const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const child = (node: Element | null, name: string) => Array.from(node?.children ?? []).find(item => item.localName === name) ?? null;
const all = (node: Document | Element, name: string) => Array.from(node.getElementsByTagNameNS('*', name));
const value = (node: Element | null, name: string, fallback = 0) => Number(node?.getAttribute(name) ?? fallback);

/** Expand table cells into DrawingML shapes so text, themes and transforms share the slide pipeline. */
export function expandTables(document: Document, styles: Document | null) {
  const make = (name: string, attributes: Record<string, string | number> = {}) => {
    const node = document.createElementNS(name.startsWith('p:') ? p : a, name);
    for (const [key, val] of Object.entries(attributes)) {
      node.setAttribute(key, String(val));
    }
    return node;
  };
  for (const frame of all(document, 'graphicFrame')) {
    const table = all(frame, 'tbl')[0];
    const transform = child(frame, 'xfrm');
    if (!table || !transform) {
      continue;
    }
    const settings = child(table, 'tblPr');
    const style = styles ? all(styles, 'tblStyle').find(item => item.getAttribute('styleId') === child(settings, 'tableStyleId')?.textContent) ?? null : null;
    const columns = Array.from(child(table, 'tblGrid')?.children ?? []).filter(item => item.localName === 'gridCol').map(item => value(item, 'w'));
    const rows = Array.from(table.children).filter(item => item.localName === 'tr');
    const group = make('p:grpSp');
    const groupProps = make('p:grpSpPr');
    const xfrm = make('a:xfrm');
    xfrm.append(...Array.from(transform.children).map(item => item.cloneNode(true)), make('a:chOff', { x: 0, y: 0 }), make('a:chExt', { cx: columns.reduce((sum, width) => sum + width, 0), cy: rows.reduce((sum, row) => sum + value(row, 'h'), 0) }));
    groupProps.append(xfrm);
    group.append(groupProps);
    let y = 0;
    rows.forEach((row, rowIndex) => {
      let x = 0;
      const cells = Array.from(row.children).filter(item => item.localName === 'tc');
      cells.forEach((cell, columnIndex) => {
        const width = columns.slice(columnIndex, columnIndex + value(cell, 'gridSpan', 1)).reduce((sum, item) => sum + item, 0);
        const height = rows.slice(rowIndex, rowIndex + value(cell, 'rowSpan', 1)).reduce((sum, item) => sum + value(item, 'h'), 0);
        if (value(cell, 'hMerge') || value(cell, 'vMerge')) {
          x += columns[columnIndex] ?? 0;
          return;
        }
        const layers = [child(style, 'wholeTbl')];
        if (value(settings, 'bandRow') && rowIndex > 0) {
          layers.push(child(style, rowIndex % 2 ? 'band1H' : 'band2H'));
        }
        for (const [flag, applies] of [['firstRow', rowIndex === 0], ['lastRow', rowIndex === rows.length - 1], ['firstCol', columnIndex === 0], ['lastCol', columnIndex === columns.length - 1]] as const) {
          if (applies && value(settings, flag)) {
            layers.push(child(style, flag));
          }
        }
        layers.reverse();
        const properties = child(cell, 'tcPr');
        const shape = make('p:sp');
        const props = make('p:spPr');
        const position = make('a:xfrm');
        position.append(make('a:off', { x, y }), make('a:ext', { cx: width, cy: height }));
        const fill = Array.from(properties?.children ?? []).find(item => ['solidFill', 'noFill', 'gradFill'].includes(item.localName)) ?? layers.map(layer => child(child(layer, 'tcStyle'), 'fill')?.firstElementChild).find(Boolean);
        props.append(position, make('a:prstGeom', { prst: 'rect' }), fill?.cloneNode(true) ?? make('a:noFill'));
        shape.append(props);
        const body = child(cell, 'txBody')?.cloneNode(true) as Element | undefined;
        if (body) {
          const bodyPr = child(body, 'bodyPr') ?? make('a:bodyPr');
          if (!bodyPr.parentElement) {
            body.prepend(bodyPr);
          }
          for (const [margin, inset, fallback] of [['marL', 'lIns', 91440], ['marR', 'rIns', 91440], ['marT', 'tIns', 45720], ['marB', 'bIns', 45720]] as const) {
            bodyPr.setAttribute(inset, String(value(properties, margin, fallback)));
          }
          bodyPr.setAttribute('anchor', properties?.getAttribute('anchor') ?? 't');
          for (const run of [...all(body, 'r'), ...all(body, 'fld')]) {
            const runProps = child(run, 'rPr') ?? make('a:rPr');
            if (!runProps.parentElement) {
              run.prepend(runProps);
            }
            for (const layer of layers) {
              const textStyle = child(layer, 'tcTxStyle');
              if (!runProps.hasAttribute('b') && textStyle?.hasAttribute('b')) {
                runProps.setAttribute('b', textStyle.getAttribute('b') === 'on' ? '1' : '0');
              }
              const color = Array.from(textStyle?.children ?? []).find(item => ['schemeClr', 'srgbClr'].includes(item.localName));
              if (color && !child(runProps, 'solidFill') && !child(runProps, 'gradFill')) {
                const solid = make('a:solidFill');
                solid.append(color.cloneNode(true));
                runProps.append(solid);
              }
            }
          }
          shape.append(body);
        }
        group.append(shape);
        for (const [edge, direct, x1, y1, x2, y2] of [['left', 'lnL', x, y, x, y + height], ['right', 'lnR', x + width, y, x + width, y + height], ['top', 'lnT', x, y, x + width, y], ['bottom', 'lnB', x, y + height, x + width, y + height]] as const) {
          const border = child(properties, direct) ?? layers.map(layer => child(child(child(child(layer, 'tcStyle'), 'tcBdr'), edge), 'ln')).find(Boolean);
          if (border) {
            const lineShape = make('p:sp');
            const lineProps = make('p:spPr');
            const lineFrame = make('a:xfrm');
            lineFrame.append(make('a:off', { x: x1, y: y1 }), make('a:ext', { cx: x2 - x1, cy: y2 - y1 }));
            const line = make('a:ln');
            for (const attr of Array.from(border.attributes)) {
              line.setAttribute(attr.name, attr.value);
            }
            line.append(...Array.from(border.children).map(item => item.cloneNode(true)));
            lineProps.append(lineFrame, make('a:prstGeom', { prst: 'line' }), make('a:noFill'), line);
            lineShape.append(lineProps);
            group.append(lineShape);
          }
        }
        x += columns[columnIndex] ?? 0;
      });
      y += value(row, 'h');
    });
    frame.replaceWith(group);
  }
}
