import { useState } from 'react';
import { Check, Paintbrush, X } from 'lucide-react';
import { Modal } from './Modal';
import './SheetColorDialog.css';

const presets = ['#dc2626', '#ea580c', '#ca8a04', '#16a34a', '#0d9488', '#2563eb', '#7c3aed', '#db2777', '#64748b'];

export function SheetColorDialog({ name, initialColor, busy, error, onApply, onClose }: {
  name: string; initialColor?: string | null; busy: boolean; error: string;
  onApply: (color: string | null) => Promise<boolean>; onClose: () => void;
}) {
  const [color, setColor] = useState(initialColor ?? null);
  return <Modal className="sheet-color-dialog" labelledBy="sheet-color-title" onClose={onClose} canClose={!busy}>
    <div className="modal-heading"><h2 id="sheet-color-title"><Paintbrush size={20} />Worksheet color</h2><button aria-label="Close worksheet color" disabled={busy} onClick={onClose}><X size={18} /></button></div>
    <p className="help">Choose a color for “{name}” in the worksheet list.</p>
    <div className="sheet-color-preview" style={{ borderLeftColor: color ?? 'transparent' }}>{name}</div>
    <div className="sheet-color-presets" role="group" aria-label="Worksheet colors">
      {presets.map(value => <button key={value} type="button" aria-label={'Color ' + value} aria-pressed={color === value} disabled={busy}
        style={{ backgroundColor: value }} onClick={() => setColor(value)}>{color === value && <Check size={18} />}</button>)}
    </div>
    <label>Custom color<input type="color" aria-label="Custom worksheet color" disabled={busy} value={color ?? '#2563eb'} onChange={event => setColor(event.target.value)} /></label>
    <button className="sheet-color-clear" type="button" aria-pressed={color === null} disabled={busy} onClick={() => setColor(null)}>No color{color === null && <Check size={16} />}</button>
    {error && <p role="alert">{error}</p>}
    <div className="modal-actions"><button disabled={busy} onClick={onClose}>Cancel</button><button className="primary" disabled={busy} onClick={async () => {
      if (await onApply(color)) {
        onClose();
      }
    }}>Save color</button></div>
  </Modal>;
}
