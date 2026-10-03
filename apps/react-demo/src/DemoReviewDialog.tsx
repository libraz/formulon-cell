import { type ReactElement, useRef } from 'react';
import type { DemoReviewDialogState, DemoUiStrings } from '../../demo-shared/index.js';
import { useDemoModalFocus } from './useDemoModalFocus.js';

interface DemoReviewDialogProps {
  dialog: DemoReviewDialogState;
  ui: DemoUiStrings;
  /** Must be referentially stable: it keys the focus-trap effect. */
  onClose: () => void;
}

export const DemoReviewDialog = ({ dialog, ui, onClose }: DemoReviewDialogProps): ReactElement => {
  const modalRef = useRef<HTMLDivElement | null>(null);
  useDemoModalFocus(modalRef, true, onClose);
  return (
    <div
      ref={modalRef}
      className="fc-tb__modal"
      role="dialog"
      aria-modal="true"
      aria-label={dialog.title}
    >
      <section className="fc-tb__modal-panel">
        <header className="fc-tb__modal-header">
          <h2>{dialog.title}</h2>
          <button type="button" className="fc-tb__modal-x" aria-label={ui.close} onClick={onClose}>
            ×
          </button>
        </header>
        <div className="fc-tb__modal-body">
          {dialog.items.length === 0 ? (
            <p className="fc-tb__modal-empty">{ui.noIssuesFound}</p>
          ) : (
            <ul className="fc-tb__modal-list">
              {dialog.items.map((item) => (
                <li key={`${item.label}-${item.detail}`}>
                  <strong>{item.label}</strong>
                  <span>{item.detail}</span>
                </li>
              ))}
            </ul>
          )}
        </div>
        <footer className="fc-tb__modal-footer">
          <button type="button" className="fc-tb__btn" onClick={onClose}>
            {ui.ok}
          </button>
        </footer>
      </section>
    </div>
  );
};
