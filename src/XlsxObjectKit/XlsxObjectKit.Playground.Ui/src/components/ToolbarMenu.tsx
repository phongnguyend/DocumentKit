import { useEffect, useId, useRef, type ReactNode } from 'react';
import { ChevronDown, type LucideIcon } from 'lucide-react';
import './ToolbarMenu.css';

export function ToolbarMenu({ label, icon: Icon, children, disabled = false }: { label: string; icon: LucideIcon; children: ReactNode; disabled?: boolean }) {
  const root = useRef<HTMLDetailsElement>(null);
  const id = useId();
  useEffect(() => {
    function outside(event: PointerEvent) {
      if (root.current && !root.current.contains(event.target as Node)) {
        root.current.open = false;
      }
    }
    document.addEventListener('pointerdown', outside);
    return () => document.removeEventListener('pointerdown', outside);
  }, []);
  useEffect(() => {
    if (disabled && root.current) {
      root.current.open = false;
    }
  }, [disabled]);
  return <details ref={root} className="toolbar-menu" name="document-actions" onKeyDown={event => {
    if (event.key === 'Escape') {
      event.stopPropagation();
      root.current!.open = false;
      root.current?.querySelector('summary')?.focus();
    }
  }}>
    <summary id={id} aria-disabled={disabled} tabIndex={disabled ? -1 : 0} onClick={event => {
      if (disabled) {
        event.preventDefault();
      }
    }} onKeyDown={event => {
      if (disabled && (event.key === 'Enter' || event.key === ' ')) {
        event.preventDefault();
      }
    }}><Icon size={16} aria-hidden="true" />{label}<ChevronDown size={13} aria-hidden="true" /></summary>
    <div className="toolbar-menu-panel" role="group" aria-labelledby={id} onClickCapture={event => {
      const button = (event.target as Element).closest('button');
      if (button && !button.disabled) {
        root.current!.open = false;
        root.current?.querySelector('summary')?.focus();
      }
    }}>{children}</div>
  </details>;
}
