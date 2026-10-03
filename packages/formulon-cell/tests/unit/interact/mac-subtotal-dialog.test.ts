import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import {
  disposeMacSubtotalDialog,
  openMacSubtotalDialog,
} from '../../../src/interact/mac-subtotal-dialog.js';
import {
  byAction,
  byId,
  createMacDataDialogFixture,
  type MacDataDialogFixture,
  pressEnter,
} from './mac-data-dialog-fixture.js';

const status = (): HTMLElement => {
  const el = document.querySelector<HTMLElement>('.fc-mac-subtotal__status');
  if (!el) throw new Error('missing status');
  return el;
};

describe('Subtotal dialog', () => {
  let fx: MacDataDialogFixture;

  beforeEach(async () => {
    fx = await createMacDataDialogFixture();
    openMacSubtotalDialog(fx.instance);
  });

  afterEach(() => {
    disposeMacSubtotalDialog(fx.instance);
    fx.dispose();
  });

  // The default one-cell selection has no subtotal column, so a run always
  // surfaces a validation error; that is how these tests observe a run.
  it('runs on Enter inside a text input', () => {
    const event = pressEnter(byId('fc-mac-subtotal-range'));
    expect(event.defaultPrevented).toBe(true);
    expect(status().dataset.state).toBe('error');
  });

  it('ignores Enter on the Cancel button', () => {
    const event = pressEnter(byAction('subtotal-cancel'));
    expect(event.defaultPrevented).toBe(false);
    expect(status().dataset.state).toBeUndefined();
  });

  it('ignores Enter inside the function select and its custom control', () => {
    pressEnter(byId('fc-mac-subtotal-function'));
    expect(status().dataset.state).toBeUndefined();
    const custom = document.querySelector('.fc-mac-subtotal .fc-select__button');
    expect(custom).not.toBeNull();
    if (custom) pressEnter(custom);
    expect(status().dataset.state).toBeUndefined();
  });

  it('ignores Enter while an IME composition is active', () => {
    pressEnter(byId('fc-mac-subtotal-range'), { isComposing: true });
    expect(status().dataset.state).toBeUndefined();
  });

  it('relabels the dialog when the locale changes', () => {
    fx.i18n.setLocale('ja');
    expect(document.querySelector('.fc-mac-subtotal .fc-fmtdlg__header')?.textContent).toBe('小計');
  });
});
