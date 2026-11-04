import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { attachConditionalDialog } from '../../../../src/interact/conditional-dialog.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../../src/store/store.js';
import { setRange } from './fixtures.js';

const fieldFor = (kind: string): HTMLInputElement => {
  const input = document
    .querySelector(`[data-range-picker="${kind}"]`)
    ?.closest('.fc-range-picker')
    ?.querySelector('input');
  if (!input) throw new Error(`missing picker field ${kind}`);
  return input;
};

describe('conditional rule dialog reference fields', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it.each([
    ['conditional-format-value-a', { kind: 'cell-value' as const, cellValueOp: '>' as const }],
    [
      'conditional-format-value-b',
      { kind: 'cell-value' as const, cellValueOp: 'between' as const },
    ],
    ['conditional-format-formula', { kind: 'formula' as const }],
  ])('%s takes a grid pick and the dialog stays open', (kind, options) => {
    const handle = attachConditionalDialog({ host, store });
    handle.open({ mode: 'new', ...options });
    const input = fieldFor(kind);
    input.focus();
    setRange(store, 1, 1, 1, 1);
    expect(input.value).toBe('=B2');
    setRange(store, 1, 1, 3, 2);
    expect(input.value).toBe('=B2:C4');
    expect(document.querySelector<HTMLElement>('.fc-conddlg')?.hidden).toBe(false);
    handle.detach();
  });
});
