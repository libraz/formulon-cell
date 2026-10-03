import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import {
  disposeMacConsolidateDialog,
  openMacConsolidateDialog,
} from '../../../src/interact/mac-consolidate-dialog.js';
import {
  byAction,
  byId,
  createMacDataDialogFixture,
  type MacDataDialogFixture,
  pressEnter,
} from './mac-data-dialog-fixture.js';

const status = (): HTMLElement => {
  const el = document.querySelector<HTMLElement>('.fc-mac-consolidate__status');
  if (!el) throw new Error('missing status');
  return el;
};

describe('Consolidate dialog', () => {
  let fx: MacDataDialogFixture;

  beforeEach(async () => {
    fx = await createMacDataDialogFixture();
    fx.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    fx.workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 2);
    fx.workbook.recalc();
    openMacConsolidateDialog(fx.instance);
  });

  afterEach(() => {
    disposeMacConsolidateDialog(fx.instance);
    fx.dispose();
  });

  it('starts with an empty destination so the selection is never the default output', () => {
    expect(byId<HTMLInputElement>('fc-mac-consolidate-destination').value).toBe('');
  });

  it('rejects a destination that overlaps a source and leaves the cells untouched', () => {
    byId<HTMLTextAreaElement>('fc-mac-consolidate-sources').value = 'A1:A2';
    byId<HTMLInputElement>('fc-mac-consolidate-destination').value = 'A1';
    byAction('consolidate-ok').click();
    expect(status().dataset.state).toBe('error');
    expect(status().textContent).toContain('overlap');
    expect(fx.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'number',
      value: 1,
    });
  });

  it('writes the aggregate to a disjoint destination on Enter', () => {
    byId<HTMLTextAreaElement>('fc-mac-consolidate-sources').value = 'A1:A2';
    const destination = byId<HTMLInputElement>('fc-mac-consolidate-destination');
    destination.value = 'C1';
    pressEnter(destination);
    expect(fx.workbook.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({
      kind: 'number',
      value: 1,
    });
    expect(fx.workbook.getValue({ sheet: 0, row: 1, col: 2 })).toEqual({
      kind: 'number',
      value: 2,
    });
  });

  it('ignores Enter on the Cancel button, the select, and during composition', () => {
    byId<HTMLTextAreaElement>('fc-mac-consolidate-sources').value = 'A1:A2';
    const destination = byId<HTMLInputElement>('fc-mac-consolidate-destination');
    destination.value = 'C1';
    pressEnter(byAction('consolidate-cancel'));
    pressEnter(byId('fc-mac-consolidate-function'));
    const custom = document.querySelector('.fc-mac-consolidate .fc-select__button');
    if (custom) pressEnter(custom);
    pressEnter(destination, { isComposing: true });
    expect(fx.workbook.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({ kind: 'blank' });
    expect(document.querySelector<HTMLElement>('.fc-mac-consolidate')?.hidden).toBe(false);
  });

  it('relabels the dialog when the locale changes', () => {
    fx.i18n.setLocale('ja');
    expect(document.querySelector('.fc-mac-consolidate .fc-fmtdlg__header')?.textContent).toBe(
      '統合',
    );
  });
});
