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

  const seedList = (): void => {
    // Headers 地域 / 商品 / 金額 and two data rows form the current region.
    const rows: (string | number)[][] = [
      ['地域', '商品', '金額'],
      ['東', 'A', 5],
      ['東', 'B', 7],
    ];
    for (const [r, row] of rows.entries()) {
      for (const [c, v] of row.entries()) {
        if (typeof v === 'number') fx.workbook.setNumber({ sheet: 0, row: r, col: c }, v);
        else fx.workbook.setText({ sheet: 0, row: r, col: c }, v);
      }
    }
    fx.store.setState((s) => ({
      ...s,
      data: {
        ...s.data,
        cells: new Map(
          [...fx.workbook.cells(0)].map((e) => [
            `0:${e.addr.row}:${e.addr.col}`,
            { value: e.value, formula: e.formula },
          ]),
        ),
      },
      selection: {
        ...s.selection,
        active: { sheet: 0, row: 0, col: 0 },
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      },
    }));
    openMacSubtotalDialog(fx.instance);
  };

  it('offers the region header names as the group column and the fields to total', () => {
    seedList();
    const group = byId<HTMLSelectElement>('fc-mac-subtotal-group-by');
    expect(Array.from(group.options, (o) => o.textContent)).toEqual(['地域', '商品', '金額']);
    expect(group.value).toBe('0');
    const fields = Array.from(
      document.querySelectorAll<HTMLInputElement>('.fc-mac-subtotal__field input'),
    );
    expect(fields.map((f) => [f.parentElement?.textContent, f.checked])).toEqual([
      ['地域', false],
      ['商品', false],
      ['金額', true],
    ]);
  });

  it('reports a missing field when none is ticked', () => {
    seedList();
    for (const f of document.querySelectorAll<HTMLInputElement>('.fc-mac-subtotal__field input')) {
      f.checked = false;
    }
    byAction('subtotal-ok').click();
    expect(status().dataset.state).toBe('error');
  });

  it('ignores Enter on the Cancel button and inside the selects', () => {
    seedList();
    const event = pressEnter(byAction('subtotal-cancel'));
    expect(event.defaultPrevented).toBe(false);
    pressEnter(byId('fc-mac-subtotal-function'));
    pressEnter(byId('fc-mac-subtotal-group-by'));
    expect(status().dataset.state).toBeUndefined();
    const custom = document.querySelector('.fc-mac-subtotal .fc-select__button');
    if (custom) pressEnter(custom);
    expect(status().dataset.state).toBeUndefined();
  });

  it('ignores Enter while an IME composition is active', () => {
    seedList();
    pressEnter(document.querySelector('.fc-mac-subtotal__field input') as Element, {
      isComposing: true,
    });
    expect(status().dataset.state).toBeUndefined();
  });

  it('Remove All closes the dialog even when the list has no subtotals', () => {
    seedList();
    byAction('subtotal-remove-all').click();
    expect(document.querySelector<HTMLElement>('.fc-mac-subtotal')?.hidden).toBe(true);
  });

  it('relabels the dialog when the locale changes', () => {
    fx.i18n.setLocale('ja');
    expect(document.querySelector('.fc-mac-subtotal .fc-fmtdlg__header')?.textContent).toBe('小計');
  });

  it('follows the desktop field order and wording', () => {
    fx.i18n.setLocale('ja');
    const rows = [...document.querySelectorAll('.fc-mac-subtotal .fc-mac-subtotal__row')].map(
      (row) => row.firstElementChild?.textContent,
    );
    expect(rows).toEqual(['グループの基準:', '集計の方法:']);
    expect(document.querySelector('.fc-mac-subtotal__fields-label')?.textContent).toBe(
      '集計するフィールド:',
    );
    const options = [...document.querySelectorAll('.fc-mac-subtotal .fc-mac-subtotal__option')].map(
      (row) => [
        row.textContent,
        row.querySelector('input')?.checked,
        row.querySelector('input')?.disabled,
      ],
    );
    expect(options).toEqual([
      ['現在の小計をすべて置き換える', true, false],
      ['グループごとに改ページを挿入する', false, true],
      ['集計行をデータの下に挿入する', true, true],
    ]);
    expect(
      [...document.querySelectorAll('.fc-mac-subtotal .fc-fmtdlg__footer button')].map(
        (b) => b.textContent,
      ),
    ).toEqual(['すべて削除', 'キャンセル', 'OK']);
  });
});
