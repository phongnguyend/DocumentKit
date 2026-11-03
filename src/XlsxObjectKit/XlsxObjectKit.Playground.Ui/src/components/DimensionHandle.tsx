import { useRef, useState } from 'react';

export function DimensionHandle({ axis, label, size, disabled, onSave }: {
  axis: 'row' | 'column'; label: string; size: number; disabled: boolean; onSave: (size: number) => void;
}) {
  const [preview, setPreview] = useState<number | null>(null);
  const gesture = useRef<{ start: number; size: number; value: number } | null>(null);
  const horizontal = axis === 'column';
  const maximum = horizontal ? 255 : 409;
  function cancel() {
    gesture.current = null;
    setPreview(null);
  }
  function update(position: number) {
    const drag = gesture.current;
    if (!drag) {
      return;
    }
    drag.value = Math.round(Math.max(1, Math.min(maximum, drag.size + (position - drag.start) * (horizontal ? 1 / 7 : 3 / 4))) * 100) / 100;
    setPreview(drag.value);
  }
  const delta = preview === null ? 0 : (preview - size) * (horizontal ? 7 : 4 / 3);
  return <span role="separator" tabIndex={disabled ? -1 : 0} aria-label={`Resize ${axis} ${label}`} aria-disabled={disabled}
    aria-orientation={horizontal ? 'vertical' : 'horizontal'} aria-valuemin={1} aria-valuemax={maximum} aria-valuenow={preview ?? size}
    title={`Drag to resize ${axis} ${label}`} className={`dimension-handle ${axis}`}
    onClick={event => event.stopPropagation()}
    onPointerDown={event => {
      event.stopPropagation();
      if (disabled || event.button !== 0) {
        return;
      }
      event.preventDefault();
      event.currentTarget.focus();
      gesture.current = { start: horizontal ? event.clientX : event.clientY, size, value: size };
      setPreview(size);
      event.currentTarget.setPointerCapture(event.pointerId);
    }} onPointerMove={event => {
      event.stopPropagation();
      update(horizontal ? event.clientX : event.clientY);
    }} onPointerUp={event => {
      event.stopPropagation();
      update(horizontal ? event.clientX : event.clientY);
      const drag = gesture.current;
      cancel();
      if (event.currentTarget.hasPointerCapture(event.pointerId)) {
        event.currentTarget.releasePointerCapture(event.pointerId);
      }
      if (drag && Math.abs(drag.value - drag.size) >= 0.01) {
        onSave(drag.value);
      }
    }} onPointerCancel={cancel} onLostPointerCapture={cancel}
    onKeyDown={event => {
      event.stopPropagation();
      if (event.key === 'Escape') {
        cancel();
      } else if (!disabled && !gesture.current && ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key)) {
        event.preventDefault();
        const value = Math.max(1, Math.min(maximum, size + (event.key === 'ArrowLeft' || event.key === 'ArrowUp' ? -1 : 1)));
        if (value !== size) {
          onSave(value);
        }
      }
    }}>
    {preview !== null && <span className="dimension-guide" style={horizontal ? { left: 4 + delta } : { top: 4 + delta }}><span role="status">{label}: {preview.toFixed(2)} {horizontal ? 'characters' : 'pt'}</span></span>}
  </span>;
}
