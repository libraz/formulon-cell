import { describe, expect, it } from 'vitest';
import { dictionaries } from '../../../src/i18n/strings.js';
import {
  backstageCardItems,
  backstageNavItems,
  createBackstageFactories,
} from '../../../src/toolbar/ribbon/backstage.js';

const ribbonText = dictionaries.en.ribbon;
const backstageText = dictionaries.en.backstage;

describe('toolbar/ribbon/backstage', () => {
  it('models the Excel-style File navigation and card actions in one shared list', () => {
    expect(backstageNavItems(backstageText, ribbonText).map((item) => item.action)).toEqual([
      'info',
      'new',
      'open',
      'save',
      'save-as',
      'print',
      'share',
      'export',
      'options',
      'close',
    ]);

    expect(backstageCardItems(backstageText, ribbonText).map((item) => item.action)).toEqual([
      'new',
      'open',
      'save',
      'save-as',
      'print',
      'page-setup',
      'edit-links',
      'share',
      'export',
      'options',
    ]);
  });

  it('renders Backstage actions from the shared model', () => {
    const docState = document.createElement('span');
    docState.textContent = 'Saved';
    const { createBackstageView } = createBackstageFactories({
      backstageText,
      ribbonText,
      shellSavedText: 'Ready',
      docName: () => 'Book1',
      docState,
    });

    const view = createBackstageView();

    expect(
      Array.from(view.querySelectorAll<HTMLElement>('.demo__backstage-navitem')).map(
        (item) => item.dataset.backstageAction,
      ),
    ).toEqual([
      'back',
      'info',
      'new',
      'open',
      'save',
      'save-as',
      'print',
      'share',
      'export',
      'options',
      'close',
    ]);
    expect(
      Array.from(view.querySelectorAll<HTMLElement>('.demo__backstage-card')).map(
        (item) => item.dataset.backstageAction,
      ),
    ).toEqual([
      'new',
      'open',
      'save',
      'save-as',
      'print',
      'page-setup',
      'edit-links',
      'share',
      'export',
      'options',
    ]);
    expect(view.textContent).toContain('Workbook Information');
    expect(view.textContent).toContain('Saved');
  });

  it('renders a dedicated Backstage Print view with command handoff attributes', () => {
    const { createBackstageView } = createBackstageFactories({
      backstageText,
      ribbonText,
      shellSavedText: 'Ready',
      docName: () => 'Book1',
      docState: null,
    });

    const view = createBackstageView('print');
    expect(view.querySelector('[data-backstage-print-preview]')).toBeTruthy();
    expect(view.textContent).toContain('Print');
    expect(view.textContent).toContain('Export to PDF');
    expect(view.textContent).toContain('Page 1');
    expect(
      Array.from(view.querySelectorAll<HTMLButtonElement>('.demo__print-action')).map(
        (button) => button.dataset.backstageAction,
      ),
    ).toEqual(['print', 'export', 'page-setup']);
    const activePrint = Array.from(
      view.querySelectorAll<HTMLElement>('.demo__backstage-navitem--active'),
    ).map((button) => button.dataset.backstageAction);
    expect(activePrint).toEqual(['print']);
  });

  it('uses a host-supplied printable document for the shared Backstage Print preview', () => {
    const { createBackstageView } = createBackstageFactories({
      backstageText,
      ribbonText,
      shellSavedText: 'Ready',
      docName: () => 'Book1',
      printPreviewHtml: () => '<!doctype html><html><body><main>Preview Cell</main></body></html>',
      docState: null,
    });

    const view = createBackstageView('print');
    const frame = view.querySelector<HTMLIFrameElement>('.demo__print-frame');

    expect(frame).toBeTruthy();
    expect(frame?.getAttribute('sandbox')).toBe('');
    expect(frame?.srcdoc).toContain('Preview Cell');
    expect(view.querySelector('.demo__print-page')).toBeNull();
  });
});
