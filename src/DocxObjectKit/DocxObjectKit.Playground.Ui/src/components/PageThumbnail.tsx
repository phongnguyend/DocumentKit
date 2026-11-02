import { useEffect, useRef, useState } from 'react';
import type { createDocxPreview } from '../renderers/docxRenderer';

type Deck = Awaited<ReturnType<typeof createDocxPreview>>;
const cache = new WeakMap<Deck, Map<number, string>>();
let queue = Promise.resolve();

export function PageThumbnail({ deck, page }: { deck: Deck; page: number }) {
  const host = useRef<HTMLSpanElement>(null);
  const [visible, setVisible] = useState(false);
  const [result, setResult] = useState<{ deck: Deck; page: number; image?: string; failed?: boolean }>();
  useEffect(() => {
    const observer = new IntersectionObserver(entries => {
      if (entries.some(entry => entry.isIntersecting)) {
        setVisible(true);
        observer.disconnect();
      }
    }, { root: host.current?.closest('.pages-panel'), rootMargin: '160px' });
    if (host.current) {
      observer.observe(host.current);
    }
    return () => observer.disconnect();
  }, []);

  useEffect(() => {
    if (!visible) {
      return;
    }
    const controller = new AbortController();
    // Serialize background previews so opening a large deck does not render every
    // page simultaneously. A new deck identity invalidates edited/reordered pages.
    queue = queue.then(async () => {
      if (controller.signal.aborted) {
        return;
      }
      try {
        let images = cache.get(deck);
        if (!images) {
          images = new Map();
          cache.set(deck, images);
        }
        let image = images.get(page);
        if (!image) {
          const canvas = document.createElement('canvas');
          const scale = 360 / Math.max(deck.width, deck.height);
          canvas.width = Math.max(1, Math.round(deck.width * scale));
          canvas.height = Math.max(1, Math.round(deck.height * scale));
          await deck.render(canvas, page, controller.signal);
          controller.signal.throwIfAborted();
          image = canvas.toDataURL('image/png');
          images.set(page, image);
        }
        setResult({ deck, page, image });
      } catch {
        if (!controller.signal.aborted) {
          setResult({ deck, page, failed: true });
        }
      }
    });
    return () => controller.abort();
  }, [deck, page, visible]);

  const current = result?.deck === deck && result.page === page ? result : undefined;
  return <span ref={host} className="page-thumbnail" style={{ aspectRatio: `${deck.width} / ${deck.height}` }} aria-hidden="true">
    {current?.image ? <img src={current.image} alt="" draggable={false} /> : <span className="thumbnail-placeholder">{current?.failed ? 'Preview unavailable' : 'Loading preview...'}</span>}
  </span>;
}
