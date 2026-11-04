import { type ReactElement, useEffect, useRef, useState } from 'react';
import { commitDemoBookName, type DemoUiStrings } from '../../demo-shared/index.js';
import { DemoGlyph } from './DemoGlyph.js';

interface DemoDocumentMenuProps {
  ui: DemoUiStrings;
  bookName: string;
  onRename: (name: string) => void;
}

/** The Mac title: book name, the document-actions chevron and its popover. */
export const DemoDocumentMenu = ({
  ui,
  bookName,
  onRename,
}: DemoDocumentMenuProps): ReactElement => {
  const [open, setOpen] = useState(false);
  const [draft, setDraft] = useState(bookName);
  const rootRef = useRef<HTMLDivElement | null>(null);
  const nameRef = useRef<HTMLInputElement | null>(null);
  const t = ui.documentMenu;

  useEffect(() => {
    if (!open) return;
    setDraft(bookName);
    // The desktop selects the name for editing as the popover opens.
    requestAnimationFrame(() => nameRef.current?.select());
    const onPointer = (event: PointerEvent): void => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    document.addEventListener('pointerdown', onPointer, true);
    return () => document.removeEventListener('pointerdown', onPointer, true);
  }, [bookName, open]);

  const commit = (): void => {
    const next = commitDemoBookName(draft, bookName);
    if (next !== bookName) onRename(next);
    setDraft(next);
  };

  return (
    <div ref={rootRef} className="fc-tb__title demo__document">
      <strong>{bookName}</strong>
      <button
        type="button"
        className="demo__document-chevron"
        aria-label={t.button}
        title={t.button}
        aria-haspopup="dialog"
        aria-expanded={open}
        onClick={() => setOpen((v) => !v)}
      />
      {open ? (
        <div className="demo__document-menu" role="dialog" aria-label={t.button}>
          <label className="demo__document-row">
            <span>{t.name}</span>
            <input
              ref={nameRef}
              type="text"
              value={draft}
              onChange={(e) => setDraft(e.currentTarget.value)}
              onBlur={commit}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  e.preventDefault();
                  commit();
                  setOpen(false);
                } else if (e.key === 'Escape') {
                  e.preventDefault();
                  setDraft(bookName);
                  setOpen(false);
                }
              }}
            />
          </label>
          <div className="demo__document-row">
            <span>{t.location}</span>
            <button type="button" className="demo__document-location" disabled>
              <DemoGlyph name="folder" />
            </button>
          </div>
          <div className="demo__document-separator" role="separator" />
          <button type="button" className="demo__document-history" disabled>
            <DemoGlyph name="versionHistory" />
            <span>{t.versionHistory}</span>
          </button>
        </div>
      ) : null}
    </div>
  );
};
