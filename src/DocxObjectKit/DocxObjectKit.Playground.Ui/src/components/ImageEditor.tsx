import { useState } from 'react';
import { Download, ImagePlus, Save, Trash2, ScanText } from 'lucide-react';
import { ImageSourceDialog } from './ImageSourceDialog';
import { ImageOcrDialog } from './ImageOcrDialog';
import { Modal } from './Modal';

export type ImageObject = { id: string; name: string; partName: string; x: number; y: number; width: number; height: number; grouped: boolean; rotation?: number };
export function ImageEditor({ image, page, loadImage, busy, onEdit, onDownload }: { image: ImageObject; page: number; loadImage: () => Promise<Blob>; busy: boolean; onEdit: (edit: object) => Promise<boolean>; onDownload: () => Promise<void> }) {
  const [bounds, setBounds] = useState({ x: image.x, y: image.y, width: image.width, height: image.height });
  const [ocr, setOcr] = useState<Blob>();
  const [loadingOcr, setLoadingOcr] = useState(false);
  const [error, setError] = useState('');
  const [confirm, setConfirm] = useState(false);
  const [replacing, setReplacing] = useState(false);
  const disabled = busy || loadingOcr;
  return <section className="text-editor image-editor"><h2>{image.name}</h2>{!image.grouped && <p className="help">Drag the selected image to move it. Drag a handle to resize. Release to save; Escape cancels.</p>}
    <button disabled={disabled} onClick={() => void onDownload().catch(cause => setError(cause.message))}><Download size={16} />Download image</button>
    <button disabled={disabled} onClick={() => {
      setLoadingOcr(true);
      setError('');
      void loadImage().then(setOcr).catch(cause => setError(cause.message)).finally(() => setLoadingOcr(false));
    }}><ScanText size={16} />{loadingOcr ? 'Preparing OCR...' : 'Recognize text (OCR)'}</button>
    {ocr && <ImageOcrDialog blob={ocr} page={page} imageName={image.name} onClose={() => setOcr(undefined)} onApply={onEdit} />}
    <button disabled={disabled} onClick={() => setReplacing(true)}><ImagePlus size={16} />Replace image</button>
    {replacing && <ImageSourceDialog initialBounds={bounds} replacing grouped={image.grouped} onClose={() => setReplacing(false)} onApply={request => onEdit({ ...request, kind: 'replace-image' })} />}
    <p className="help">Replacement keeps the existing position, crop and rotation.</p>
    {image.grouped ? <p className="help">This image belongs to a group. Position and size editing is not available yet.</p> : <><div className="image-bounds">{(['x', 'y', 'width', 'height'] as const).map(key => <label key={key}>{key.toUpperCase()} (pt)<input type="number" aria-label={`Image ${key}`} value={bounds[key]} min={key === 'width' || key === 'height' ? 0.01 : undefined} step="0.1" onChange={event => setBounds(previous => ({ ...previous, [key]: Number(event.target.value) }))} disabled={disabled} /></label>)}</div><button disabled={disabled || bounds.width <= 0 || bounds.height <= 0 || !Object.values(bounds).every(Number.isFinite)} onClick={() => void onEdit({ kind: 'resize-image', ...bounds })}><Save size={16} />Save position and size</button></>}
    <button className="danger" disabled={disabled} onClick={() => setConfirm(true)}><Trash2 size={16} />Delete image</button>
    {error && <p role="alert">{error}</p>}
    {confirm && <Modal labelledBy="delete-image-title" className="confirm-modal" onClose={() => setConfirm(false)}><h2 id="delete-image-title">Delete image?</h2><p>Remove this image from the page. You can undo this change.</p><div className="modal-actions"><button onClick={() => setConfirm(false)}>Cancel</button><button className="danger" onClick={() => { setConfirm(false); void onEdit({ kind: 'delete-image' }); }}>Delete image</button></div></Modal>}
  </section>;
}
