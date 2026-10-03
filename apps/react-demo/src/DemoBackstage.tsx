import type { SpreadsheetInstance } from '@libraz/formulon-cell';
import { type ReactElement, useMemo } from 'react';
import {
  buildDemoBackstageCards,
  buildDemoBackstageNav,
  buildDemoPrintPreviewModel,
  DEMO_PRINT_PREVIEW_LINES,
  type DemoBackstageAction,
  type DemoUiStrings,
  isDemoBackstageActionDisabled,
} from '../../demo-shared/index.js';
import { DemoIcon } from './DemoIcon.js';

interface DemoBackstageProps {
  ui: DemoUiStrings;
  instance: SpreadsheetInstance | null;
  bookName: string;
  action: DemoBackstageAction;
  onAction: (action: DemoBackstageAction) => void;
}

export const DemoBackstage = ({
  ui,
  instance,
  bookName,
  action,
  onAction,
}: DemoBackstageProps): ReactElement => {
  const nav = useMemo(() => buildDemoBackstageNav(ui, action), [action, ui]);
  const cards = useMemo(() => buildDemoBackstageCards(ui), [ui]);
  const printPreview = useMemo(() => {
    void action;
    return buildDemoPrintPreviewModel(ui, instance, bookName);
  }, [action, bookName, instance, ui]);
  return (
    <div className="fc-tb__backstage" role="dialog" aria-label={ui.file}>
      <nav className="fc-tb__backstage-nav" aria-label={ui.file}>
        <strong>{ui.file}</strong>
        {nav.map((item) => (
          <button
            key={item.action}
            type="button"
            className={`fc-tb__backstage-navitem${
              item.active ? ' fc-tb__backstage-navitem--active' : ''
            }`}
            onClick={() => onAction(item.action)}
            disabled={isDemoBackstageActionDisabled(item.action, instance)}
          >
            {item.label}
          </button>
        ))}
      </nav>
      <div className="fc-tb__backstage-main">
        <div className="fc-tb__backstage-title">
          <span className="fc-tb__backstage-xl" aria-hidden="true">
            <DemoIcon name="app" />
          </span>
          <div>
            <h1>{bookName}</h1>
            <p>{ui.backstageSub}</p>
          </div>
        </div>
        {action === 'print' ? (
          <div className="fc-tb__print-preview" data-demo-print-preview>
            <section className="fc-tb__print-settings" aria-label={ui.printSettings}>
              <h2>{printPreview.title}</h2>
              <p>{printPreview.subtitle}</p>
              <button
                type="button"
                className="fc-tb__print-action fc-tb__print-action--primary"
                onClick={() => instance?.print('print')}
                disabled={!instance}
              >
                {printPreview.printLabel}
              </button>
              <button
                type="button"
                className="fc-tb__print-action"
                onClick={() => instance?.print('pdf')}
                disabled={!instance}
              >
                {printPreview.pdfLabel}
              </button>
              <button
                type="button"
                className="fc-tb__print-action"
                onClick={() => instance?.openPageSetup()}
                disabled={!instance}
              >
                {printPreview.pageSetupLabel}
              </button>
              <dl className="fc-tb__print-meta">
                {printPreview.settings.map((row) => (
                  <div key={row.label}>
                    <dt>{row.label}</dt>
                    <dd>{row.value}</dd>
                  </div>
                ))}
              </dl>
            </section>
            <section className="fc-tb__print-paper" aria-label={printPreview.previewTitle}>
              {printPreview.previewHtml ? (
                <iframe
                  className="fc-tb__print-frame"
                  title={printPreview.previewTitle}
                  sandbox=""
                  srcDoc={printPreview.previewHtml}
                />
              ) : (
                <div className="fc-tb__print-page">
                  <strong>{printPreview.previewTitle}</strong>
                  <div aria-hidden="true" className="fc-tb__print-sheet-lines">
                    {DEMO_PRINT_PREVIEW_LINES.map((line) => (
                      <span key={line} />
                    ))}
                  </div>
                </div>
              )}
              <p>{printPreview.previewHint}</p>
            </section>
          </div>
        ) : (
          <div className="fc-tb__backstage-grid">
            {cards.map((item) => (
              <button
                key={item.action}
                type="button"
                className="fc-tb__backstage-card"
                onClick={() => onAction(item.action)}
                disabled={isDemoBackstageActionDisabled(item.action, instance)}
              >
                <strong>{item.label}</strong>
                <span>{item.desc}</span>
              </button>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};
