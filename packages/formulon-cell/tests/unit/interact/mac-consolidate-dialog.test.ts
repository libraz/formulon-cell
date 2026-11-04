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

  it('writes the result at the active cell, never over the source', () => {
    fx.store.setState((s) => ({
      ...s,
      selection: {
        ...s.selection,
        active: { sheet: 0, row: 0, col: 2 },
        range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
      },
    }));
    openMacConsolidateDialog(fx.instance);
    byId<HTMLInputElement>('fc-mac-consolidate-sources').value = 'A1:A2';
    byAction('consolidate-ok').click();
    expect(fx.workbook.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({
      kind: 'number',
      value: 1,
    });
    expect(fx.workbook.getValue({ sheet: 0, row: 1, col: 2 })).toEqual({
      kind: 'number',
      value: 2,
    });
  });

  it('rejects a result that would land on a source and leaves the cells untouched', () => {
    byId<HTMLInputElement>('fc-mac-consolidate-sources').value = 'A1:A2';
    byAction('consolidate-ok').click();
    expect(status().dataset.state).toBe('error');
    expect(status().textContent).toContain('overlap');
    expect(fx.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'number',
      value: 1,
    });
  });

  it('adds typed references to the list and removes the selected one', () => {
    const input = byId<HTMLInputElement>('fc-mac-consolidate-sources');
    input.value = 'A1:A2';
    byAction('consolidate-add').click();
    input.value = 'C1:C2';
    byAction('consolidate-add').click();
    const items = () =>
      Array.from(
        document.querySelectorAll('.fc-mac-consolidate__reference'),
        (el) => el.textContent,
      );
    expect(items()).toEqual(['A1:A2', 'C1:C2']);
    expect(input.value).toBe('');
    document.querySelectorAll<HTMLElement>('.fc-mac-consolidate__reference')[0]?.click();
    byAction('consolidate-remove').click();
    expect(items()).toEqual(['C1:C2']);
  });

  it('offers the top row and left column label options, enabled and unchecked', () => {
    const options = Array.from(
      document.querySelectorAll<HTMLInputElement>('.fc-mac-consolidate__criteria input'),
    );
    expect(options).toHaveLength(2);
    expect(options.map((o) => [o.checked, o.disabled])).toEqual([
      [false, false],
      [false, false],
    ]);
  });

  it('ignores Enter on the Cancel button, the select, and during composition', () => {
    fx.store.setState((s) => ({
      ...s,
      selection: {
        ...s.selection,
        active: { sheet: 0, row: 0, col: 2 },
        range: { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
      },
    }));
    openMacConsolidateDialog(fx.instance);
    const input = byId<HTMLInputElement>('fc-mac-consolidate-sources');
    input.value = 'A1:A2';
    pressEnter(byAction('consolidate-cancel'));
    pressEnter(byId('fc-mac-consolidate-function'));
    const custom = document.querySelector('.fc-mac-consolidate .fc-select__button');
    if (custom) pressEnter(custom);
    pressEnter(input, { isComposing: true });
    expect(fx.workbook.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({ kind: 'blank' });
    expect(document.querySelector<HTMLElement>('.fc-mac-consolidate')?.hidden).toBe(false);
    pressEnter(input);
    expect(fx.workbook.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({
      kind: 'number',
      value: 1,
    });
  });

  it('relabels the dialog when the locale changes', () => {
    fx.i18n.setLocale('ja');
    expect(document.querySelector('.fc-mac-consolidate .fc-fmtdlg__header')?.textContent).toBe(
      '統合の設定',
    );
  });

  it('follows the desktop 統合の設定 order and wording', () => {
    fx.i18n.setLocale('ja');
    const panel = document.querySelector<HTMLElement>('.fc-mac-consolidate');
    expect(panel?.querySelector('.fc-fmtdlg__header')?.textContent).toBe('統合の設定');
    const rows = [...(panel?.querySelectorAll('.fc-mac-consolidate__row') ?? [])].map(
      (row) => row.firstElementChild?.textContent,
    );
    expect(rows.slice(0, 3)).toEqual(['集計の方法:', '統合元範囲:', '統合元:']);
    expect(panel?.querySelector('.fc-mac-consolidate__criteria-heading')?.textContent).toBe(
      '統合の基準:',
    );
    expect(
      [
        ...(panel?.querySelectorAll('.fc-mac-consolidate__criteria .fc-mac-consolidate__option') ??
          []),
      ].map((el) => el.textContent),
    ).toEqual(['上端行', '左端列']);
    const options = [...(panel?.querySelectorAll<HTMLElement>('.fc-mac-consolidate__option') ?? [])]
      .filter((el) => !el.hidden)
      .map((el) => el.textContent);
    expect(options[0]).toBe('統合元データとリンクする');
    expect(byAction('consolidate-cancel').textContent).toBe('閉じる');
  });
});
