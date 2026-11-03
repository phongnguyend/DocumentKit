import { useState } from 'react';
import { Plus, Trash2, X } from 'lucide-react';
import { Modal } from './Modal';
import type { Sheet } from '../types';

export const comparisons: Record<string, string> = {
  greaterThan: 'Greater than', greaterThanOrEqual: 'Greater than or equal to', lessThan: 'Less than',
  lessThanOrEqual: 'Less than or equal to', equal: 'Equal to', notEqual: 'Not equal to', between: 'Between', notBetween: 'Not between',
};

export function ConditionalFormatDialog({ sheet, initialRange, busy, onEdit, onClose }: {
  sheet: Sheet; initialRange: string; busy: boolean; onEdit: (request: object) => Promise<boolean>; onClose: () => void;
}) {
  const [range, setRange] = useState(initialRange);
  const [operator, setOperator] = useState('greaterThan');
  const [value, setValue] = useState('0');
  const [second, setSecond] = useState('100');
  const [color, setColor] = useState('#9c0006');
  const [background, setBackground] = useState('#ffc7ce');
  const [failed, setFailed] = useState(false);
  const interval = operator === 'between' || operator === 'notBetween';
  return <Modal labelledBy="conditional-title" onClose={onClose} canClose={!busy}>
    <div className="modal-heading"><h2 id="conditional-title">Conditional formatting</h2><button aria-label="Close conditional formatting" disabled={busy} onClick={onClose}><X size={18} /></button></div>
    <p className="help">Highlight numeric values and calculated numeric results. Rules apply in priority order without changing the underlying cell formatting.</p>
    <form onSubmit={async event => {
      event.preventDefault();
      setFailed(!await onEdit({ action: 'add-conditional-rule', range, operator, value: Number(value), secondValue: Number(second), color, background }));
    }}>
      <label>Applies to<input aria-label="Applies to" required value={range} disabled={busy} onChange={e => setRange(e.target.value)} placeholder="A1:D20" /></label>
      <label>Condition<select aria-label="Condition" disabled={busy} value={operator} onChange={e => setOperator(e.target.value)}>{Object.entries(comparisons).map(([key, label]) => <option key={key} value={key}>{label}</option>)}</select></label>
      <label>{interval ? 'Minimum' : 'Value'}<input aria-label="Rule value" type="number" step="any" required value={value} disabled={busy} onChange={e => setValue(e.target.value)} /></label>
      {interval && <label>Maximum<input aria-label="Maximum" type="number" step="any" min={value} required value={second} disabled={busy} onChange={e => setSecond(e.target.value)} /></label>}
      <div className="modal-actions"><label>Text color<input aria-label="Rule text color" type="color" value={color} disabled={busy} onChange={e => setColor(e.target.value)} /></label><label>Fill color<input aria-label="Rule fill color" type="color" value={background} disabled={busy} onChange={e => setBackground(e.target.value)} /></label><span style={{ color, background, padding: '8px 16px', borderRadius: 6 }}>123.45</span></div>
      <button className="primary" type="submit" disabled={busy}><Plus size={16} />Add rule</button>
    </form>
    {failed && <p role="alert">Could not add the rule. Check the range and values; details appear in the workbook error message.</p>}
    <h3>Worksheet rules ({sheet.conditionalRules?.length ?? 0})</h3>
    <p className="help">Higher-priority matches take precedence. To change a rule, remove it and add its replacement. Advanced imported rules are preserved but may not be previewed.</p>
    {(sheet.conditionalRules ?? []).slice().sort((a, b) => a.priority - b.priority).map(rule => <div key={rule.id} className="modal-actions">
      <span style={{ flex: 1 }}><strong>{rule.range}</strong><br />{comparisons[rule.operator] ?? rule.operator} {rule.value}{rule.secondValue != null ? ` – ${rule.secondValue}` : ''}{!rule.supported ? ' (preview unsupported)' : ''}</span>
      <button disabled={busy} aria-label={`Remove rule ${rule.id + 1}`} onClick={() => void onEdit({ action: 'remove-conditional-rule', ruleId: rule.id })}><Trash2 size={16} />Remove</button>
    </div>)}
    <div className="modal-actions"><button disabled={busy} onClick={onClose}>Done</button></div>
  </Modal>;
}
