import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { setMarginPreset } from '../../../src/commands/page-setup.js';
import { attachPageSetupDialog } from '../../../src/interact/page-setup-dialog.js';
import {
  createSpreadsheetStore,
  defaultPageSetup,
  getPageSetup,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const dialog = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-pgsetup');
const selects = (): HTMLSelectElement[] =>
  Array.from(document.querySelectorAll<HTMLSelectElement>('.fc-pgsetup__select'));
const marginInputs = (): HTMLInputElement[] =>
  Array.from(document.querySelectorAll<HTMLInputElement>('.fc-pgsetup__margins input'));
const buttons = (): HTMLButtonElement[] =>
  Array.from(document.querySelectorAll<HTMLButtonElement>('.fc-pgsetup .fc-fmtdlg__btn'));

describe('attachPageSetupDialog', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
    store = createSpreadsheetStore();
  });

  afterEach(() => {
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  it('mounts a hidden overlay and shows the dialog on open', () => {
    const handle = attachPageSetupDialog({ host, store });
    expect(dialog()?.hidden).toBe(true);
    handle.open();
    expect(dialog()?.hidden).toBe(false);
    handle.detach();
  });

  it('hydrates inputs from defaults and reflects the orientation/paper select values', () => {
    const handle = attachPageSetupDialog({ host, store });
    handle.open();
    const [orientSelect, paperSelect] = selects();
    expect(orientSelect?.value).toBe(defaultPageSetup().orientation);
    expect(paperSelect?.value).toBe(defaultPageSetup().paperSize);
    expect(orientSelect?.options.length).toBe(2);
    expect(paperSelect?.options.length).toBe(6);
    handle.detach();
  });

  it('OK with a flipped orientation/paper persists via mutators.setPageSetup', () => {
    const handle = attachPageSetupDialog({ host, store });
    handle.open();
    const [orientSelect, paperSelect] = selects();
    if (!orientSelect || !paperSelect) throw new Error('selects missing');
    orientSelect.value = 'landscape';
    paperSelect.value = 'letter';
    const ok = buttons().find((b) => b.classList.contains('fc-fmtdlg__btn--primary'));
    ok?.click();
    const setup = getPageSetup(store.getState(), 0);
    expect(setup.orientation).toBe('landscape');
    expect(setup.paperSize).toBe('letter');
    expect(dialog()?.hidden).toBe(true);
    handle.detach();
  });

  it('Cancel does not mutate the page setup', () => {
    // Seed with a known value so the dialog can be canceled and we can confirm
    // the slice didn't move.
    setMarginPreset(store, 0, 'narrow');
    const before = getPageSetup(store.getState(), 0);
    const handle = attachPageSetupDialog({ host, store });
    handle.open();
    // Mutate inputs in the dialog.
    const [orientSelect] = selects();
    if (orientSelect) orientSelect.value = 'landscape';
    const inputs = marginInputs();
    if (inputs[0]) inputs[0].value = '9';
    const cancel = buttons().find((b) => !b.classList.contains('fc-fmtdlg__btn--primary'));
    cancel?.click();
    const after = getPageSetup(store.getState(), 0);
    expect(after).toEqual(before);
    expect(dialog()?.hidden).toBe(true);
    handle.detach();
  });

  it('Escape closes the overlay without mutating', () => {
    const before = getPageSetup(store.getState(), 0);
    const handle = attachPageSetupDialog({ host, store });
    handle.open();
    const overlay = dialog();
    overlay?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(overlay?.hidden).toBe(true);
    expect(getPageSetup(store.getState(), 0)).toEqual(before);
    handle.detach();
  });

  it('OK after editing margin inputs writes the updated margins to the slice', () => {
    const handle = attachPageSetupDialog({ host, store });
    handle.open();
    const inputs = marginInputs();
    if (inputs.length !== 4) throw new Error('expected 4 margin inputs');
    const [top, right, bottom, left] = inputs as [
      HTMLInputElement,
      HTMLInputElement,
      HTMLInputElement,
      HTMLInputElement,
    ];
    top.value = '0.5';
    right.value = '0.4';
    bottom.value = '0.3';
    left.value = '0.2';
    const ok = buttons().find((b) => b.classList.contains('fc-fmtdlg__btn--primary'));
    ok?.click();
    const setup = getPageSetup(store.getState(), 0);
    expect(setup.margins).toEqual({ top: 0.5, right: 0.4, bottom: 0.3, left: 0.2 });
    handle.detach();
  });

  it('detach removes the overlay node from the DOM', () => {
    const handle = attachPageSetupDialog({ host, store });
    handle.detach();
    expect(dialog()).toBeNull();
  });
});
