import { useEffect, useRef, useState, type PointerEvent } from 'react';
type TransformObject = { id: string; x: number; y: number; width: number; height: number; rotation?: number };

type Bounds = { x: number; y: number; width: number; height: number };
type Gesture = { pointer: number; startX: number; startY: number; scaleX: number; scaleY: number; handle: string; original: Bounds; next: Bounds; target: HTMLElement };
export function ImageTransform({ image, slideWidth, slideHeight, busy, onSave, kind = 'image' }: { kind?: 'image' | 'text'; image: TransformObject; slideWidth: number; slideHeight: number; busy: boolean; onSave: (bounds: Bounds) => Promise<boolean> }) {
  const [preview, setPreview] = useState<Bounds>(image);
  const gesture = useRef<Gesture | null>(null);
  useEffect(() => {
    setPreview(image);
  }, [image]);
  function cancel() {
    const active = gesture.current;
    gesture.current = null;
    if (active?.target.hasPointerCapture(active.pointer)) {
      active.target.releasePointerCapture(active.pointer);
    }
    setPreview(image);
  }
  useEffect(() => {
    function escape(event: KeyboardEvent) {
      if (event.key === 'Escape' && gesture.current) {
        event.preventDefault();
        cancel();
      }
    }
    window.addEventListener('keydown', escape);
    return () => window.removeEventListener('keydown', escape);
  }, [image]);
  function start(event: PointerEvent<HTMLElement>, handle: string) {
    if (busy || event.button !== 0 || gesture.current) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    const rect = event.currentTarget.closest('.slide-stage')!.getBoundingClientRect();
    gesture.current = { pointer: event.pointerId, startX: event.clientX, startY: event.clientY, scaleX: slideWidth / rect.width, scaleY: slideHeight / rect.height, handle, original: image, next: image, target: event.currentTarget };
    event.currentTarget.setPointerCapture(event.pointerId);
  }
  function move(event: PointerEvent<HTMLElement>) {
    const active = gesture.current;
    if (!active || active.pointer !== event.pointerId) {
      return;
    }
    const dx = (event.clientX - active.startX) * active.scaleX;
    const dy = (event.clientY - active.startY) * active.scaleY;
    const original = active.original;
    let next: Bounds;
    if (active.handle === 'move') {
      next = { ...original, x: original.x + dx, y: original.y + dy };
    } else {
      const angle = (image.rotation ?? 0) * Math.PI / 180;
      const cos = Math.cos(angle), sin = Math.sin(angle);
      const localX = dx * cos + dy * sin, localY = -dx * sin + dy * cos;
      const horizontal = active.handle.includes('e') ? 1 : active.handle.includes('w') ? -1 : 0;
      const vertical = active.handle.includes('s') ? 1 : active.handle.includes('n') ? -1 : 0;
      const width = Math.max(1, original.width + horizontal * localX);
      const height = Math.max(1, original.height + vertical * localY);
      const cx = horizontal * (width - original.width) / 2, cy = vertical * (height - original.height) / 2;
      next = { width, height, x: original.x + original.width / 2 + cx * cos - cy * sin - width / 2, y: original.y + original.height / 2 + cx * sin + cy * cos - height / 2 };
    }
    active.next = next;
    setPreview(next);
  }
  function finish(event: PointerEvent<HTMLElement>) {
    const active = gesture.current;
    if (!active || active.pointer !== event.pointerId) {
      return;
    }
    move(event);
    gesture.current = null;
    active.target.releasePointerCapture(active.pointer);
    const next = Object.fromEntries(['x', 'y', 'width', 'height'].map(key => [key, Math.round(active.next[key as keyof Bounds] * 1000) / 1000])) as Bounds;
    if (['x', 'y', 'width', 'height'].some(key => Math.abs(next[key as keyof Bounds] - active.original[key as keyof Bounds]) > .01)) {
      void onSave(next).then(success => {
        if (!success) {
          setPreview(image);
        }
      });
    } else {
      setPreview(image);
    }
  }
  return <div className="image-transform" aria-label={`Move selected ${kind}`} role="group" style={{ left: `${preview.x / slideWidth * 100}%`, top: `${preview.y / slideHeight * 100}%`, width: `${preview.width / slideWidth * 100}%`, height: `${preview.height / slideHeight * 100}%`, transform: `rotate(${image.rotation ?? 0}deg)`, cursor: busy ? 'wait' : 'move' }} onClick={event => event.stopPropagation()} onPointerDown={event => start(event, 'move')} onPointerMove={move} onPointerUp={finish} onPointerCancel={cancel} onLostPointerCapture={() => {
    if (gesture.current) {
      cancel();
    }
  }}>
    {['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'].map(handle => <button key={handle} type="button" className={`image-resize-handle handle-${handle}`} aria-label={`Resize ${kind} ${handle}`} disabled={busy} onPointerDown={event => start(event, handle)} />)}
  </div>;
}
