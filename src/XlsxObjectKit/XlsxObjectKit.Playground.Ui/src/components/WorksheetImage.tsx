import { useEffect, useRef, useState } from 'react';
import type { SheetImage } from '../types';

export function WorksheetImage({ image, selected, busy, offsetX, offsetY, onSelect, onSave, onDelete }: {
  image: SheetImage; selected: boolean; busy: boolean; offsetX: number; offsetY: number;
  onSelect: () => void; onSave: (bounds: object) => void; onDelete: () => void;
}) {
  const [bounds, setBounds] = useState(image);
  const gesture = useRef<{ x: number; y: number; handle: string; original: SheetImage; current: SheetImage } | null>(null);
  useEffect(() => {
    setBounds(image);
  }, [image]);
  function cancel() {
    if (gesture.current) {
      setBounds(gesture.current.original);
      gesture.current = null;
    }
  }
  return <div role="button" tabIndex={0} aria-label={`Select image ${image.name}`} aria-pressed={selected}
    className={'worksheet-image' + (selected ? ' selected' : '')}
    style={{ left: bounds.x + offsetX, top: bounds.y + offsetY, width: bounds.width, height: bounds.height }}
    onClick={event => event.stopPropagation()} onDoubleClick={event => event.stopPropagation()}
    onKeyDown={event => {
      event.stopPropagation();
      if (event.key === 'Escape') {
        cancel();
      } else if (!busy && selected && event.key === 'Delete') {
        event.preventDefault();
        onDelete();
      } else if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        onSelect();
      }
    }}
    onPointerDown={event => {
      event.stopPropagation();
      if (busy || event.button !== 0) {
        return;
      }
      event.preventDefault();
      onSelect();
      event.currentTarget.focus();
      if (image.editable) {
        gesture.current = { x: event.clientX, y: event.clientY, handle: (event.target as HTMLElement).dataset.handle ?? '', original: image, current: image };
        event.currentTarget.setPointerCapture(event.pointerId);
      }
    }} onPointerMove={event => {
      event.stopPropagation();
      const drag = gesture.current;
      if (!drag) {
        return;
      }
      const dx = event.clientX - drag.x;
      const dy = event.clientY - drag.y;
      const next = { ...drag.original };
      if (!drag.handle) {
        next.x = Math.max(0, next.x + dx);
        next.y = Math.max(0, next.y + dy);
      } else {
        if (drag.handle.includes('e')) {
          next.width = Math.max(8, next.width + dx);
        }
        if (drag.handle.includes('s')) {
          next.height = Math.max(8, next.height + dy);
        }
        if (drag.handle.includes('w')) {
          next.x = Math.max(0, Math.min(next.x + dx, next.x + next.width - 8));
          next.width += drag.original.x - next.x;
        }
        if (drag.handle.includes('n')) {
          next.y = Math.max(0, Math.min(next.y + dy, next.y + next.height - 8));
          next.height += drag.original.y - next.y;
        }
      }
      drag.current = next;
      setBounds(next);
    }} onPointerUp={event => {
      event.stopPropagation();
      const drag = gesture.current;
      gesture.current = null;
      if (event.currentTarget.hasPointerCapture(event.pointerId)) {
        event.currentTarget.releasePointerCapture(event.pointerId);
      }
      if (drag && ['x', 'y', 'width', 'height'].some(key => drag.original[key as keyof SheetImage] !== drag.current[key as keyof SheetImage])) {
        const { x, y, width, height } = drag.current;
        setBounds(image);
        onSave({ x, y, width, height });
      }
    }} onPointerCancel={cancel} onLostPointerCapture={cancel}>
    <img src={image.dataUrl} alt={image.name} draggable={false} />
    {selected && image.editable && ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'].map(handle => <span key={handle} data-handle={handle} className={'worksheet-image-handle ' + handle} aria-label={`Resize image ${handle}`} />)}
  </div>;
}
