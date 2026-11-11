import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { defaultTableOverlay } from '../../../src/commands/format-as-table.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings/en.js';
import { attachMacSlicerDialog } from '../../../src/interact/mac-slicer-dialog.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const workbook = {
  getTables: () => [
    {
      name: 'Sales',
      displayName: 'Sales',
      ref: 'A1:B3',
      sheetIndex: 0,
      columns: ['Region', 'Amount'],
    },
    { name: 'Costs', displayName: 'Costs', ref: 'D1:E3', sheetIndex: 0, columns: ['Item'] },
  ],
} as unknown as WorkbookHandle;

const storeWithTable = (
  name: string,
  source: 'engine' | 'session',
  range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
): SpreadsheetStore => {
  const store = createSpreadsheetStore();
  mutators.upsertTableOverlay(store, { ...defaultTableOverlay('t1', range), source, name });
  return store;
};

describe('Insert Slicer dialog targeting', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  const ok = (): HTMLButtonElement => {
    const found = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    if (!found) throw new Error('missing OK');
    return found;
  };
  const errorText = (): string | undefined =>
    document.querySelector('[data-mac-slicer-error]')?.textContent ?? undefined;
  const fieldNames = (): string[] =>
    Array.from(
      document.querySelectorAll('.fc-macslicerdlg__field'),
      (row) => row.textContent ?? '',
    );

  it('lists the columns of the active engine table, not the first table', () => {
    const store = storeWithTable('Costs', 'engine', { sheet: 0, r0: 0, c0: 3, r1: 2, c1: 4 });
    mutators.setActive(store, { sheet: 0, row: 1, col: 3 });
    const handle = attachMacSlicerDialog({
      host,
      store,
      getWb: () => workbook,
      strings: en,
      onAdd: () => null,
    });
    handle.open();
    expect(fieldNames()).toEqual(['Item']);
    handle.detach();
  });

  it.each([
    ['the active cell is outside every table', createSpreadsheetStore()],
    ['the active table is session-only', storeWithTable('Sales', 'session')],
  ])('stays unavailable when %s', (_label, store) => {
    const handle = attachMacSlicerDialog({
      host,
      store,
      getWb: () => workbook,
      strings: en,
      onAdd: () => null,
    });
    handle.open();
    expect(fieldNames()).toEqual([]);
    expect(ok().disabled).toBe(true);
    expect(errorText()).toBe(en.mac.slicer.noTable);
    handle.detach();
  });

  it('stays unavailable when opened for a PivotTable', () => {
    const handle = attachMacSlicerDialog({
      host,
      store: storeWithTable('Sales', 'engine'),
      getWb: () => workbook,
      strings: en,
      onAdd: () => null,
    });
    handle.open({ pivot: true });
    expect(fieldNames()).toEqual([]);
    expect(errorText()).toBe(en.mac.slicer.noTable);
    handle.open();
    expect(fieldNames()).toEqual(['Region', 'Amount']);
    handle.detach();
  });

  it('adds every checked field in one call under the invoking command id', () => {
    const calls: { inputs: unknown; commandId: string }[] = [];
    const handle = attachMacSlicerDialog({
      host,
      store: storeWithTable('Sales', 'engine'),
      getWb: () => workbook,
      strings: en,
      onAdd: (inputs, commandId) => {
        calls.push({ inputs, commandId });
        return [];
      },
    });
    handle.open({ commandId: 'tableInsertSlicer' });
    for (const box of document.querySelectorAll<HTMLInputElement>('input[type="checkbox"]')) {
      box.click();
    }
    ok().click();
    expect(calls).toEqual([
      {
        inputs: [
          { tableName: 'Sales', column: 'Region' },
          { tableName: 'Sales', column: 'Amount' },
        ],
        commandId: 'tableInsertSlicer',
      },
    ]);
    handle.detach();
  });

  it('keeps the dialog open with an error when the batch is refused', () => {
    const handle = attachMacSlicerDialog({
      host,
      store: storeWithTable('Sales', 'engine'),
      getWb: () => workbook,
      strings: en,
      onAdd: () => null,
    });
    handle.open();
    document.querySelector<HTMLInputElement>('input[type="checkbox"]')?.click();
    ok().click();
    expect(errorText()).toBe(en.mac.slicer.insertFailed);
    expect(document.querySelector<HTMLElement>('.fc-macslicerdlg')?.hidden).toBe(false);
    handle.detach();
  });
});
