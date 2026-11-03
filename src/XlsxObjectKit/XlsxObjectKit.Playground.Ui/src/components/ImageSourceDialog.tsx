import { useState } from 'react';
import { ImagePlus, X } from 'lucide-react';
import { ImagePicker } from './ImagePicker';
import { Modal } from './Modal';
import { prepareImageForXlsx } from '../prepareImage';

export type ImageBounds = { x: number; y: number; width: number; height: number };
export function ImageSourceDialog({ initialBounds, replacing = false, grouped = false, onClose, onApply }: { initialBounds: ImageBounds; replacing?: boolean; grouped?: boolean; onClose: () => void; onApply: (request: object) => Promise<boolean> }) {
  const [image, setImage] = useState<File>();
  const [bounds, setBounds] = useState<ImageBounds>(() => ({ x: initialBounds.x, y: initialBounds.y, width: initialBounds.width, height: initialBounds.height }));
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  async function apply() {
    if (!image || busy) {
      return;
    }
    setBusy(true);
    setError('');
    try {
      const prepared = await prepareImageForXlsx(image);
      const imageBase64 = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result).split(',')[1]);
        reader.onerror = () => reject(new Error('Could not read the image.'));
        reader.readAsDataURL(prepared);
      });
      const success = await onApply({ imageBase64, ...bounds });
      if (success) {
        onClose();
      } else {
        setError('The image could not be saved. See the document error for details.');
      }
    } catch (cause) {
      setError((cause as Error).message);
    } finally {
      setBusy(false);
    }
  }
  return <Modal labelledBy="image-source-title" className="image-source-dialog" onClose={onClose} canClose={!busy}><div className="modal-heading"><h2 id="image-source-title">{replacing ? 'Replace image' : 'Add image'}</h2><button aria-label="Close image picker" disabled={busy} onClick={onClose}><X size={18} /></button></div>
    <fieldset disabled={busy} className="image-source-controls"><ImagePicker image={image} width={bounds.width} height={bounds.height} onChange={setImage} onUseSize={grouped || replacing ? undefined : (width, height) => setBounds(previous => ({ ...previous, width, height }))} />
    {!grouped && !replacing && <div className="image-bounds">{(['x', 'y', 'width', 'height'] as const).map(key => <label key={key}>{key.toUpperCase()} (px)<input aria-label={`Placement ${key}`} type="number" step="0.1" value={bounds[key]} onChange={event => setBounds(previous => ({ ...previous, [key]: Number(event.target.value) }))} /></label>)}</div>}</fieldset>
    <p className="help">{replacing ? 'Keeps the existing crop and rotation. Drawings use the placement aspect ratio.' : 'Upload, draw, or generate a QR code. Drawings use the placement aspect ratio.'}</p>
    {error && <p role="alert">{error}</p>}<div className="modal-actions"><button disabled={busy} onClick={onClose}>Cancel</button><button className="primary" disabled={busy || !image || !Object.values(bounds).every(Number.isFinite) || (!grouped && (bounds.width <= 0 || bounds.height <= 0))} onClick={() => void apply()}><ImagePlus size={16} />{busy ? 'Saving...' : replacing ? 'Replace image' : 'Insert image'}</button></div>
  </Modal>;
}
