import { beforeEach, describe, expect, it } from 'vitest';
import type { CellValue } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { attachPivotTableDialog } from '../../../src/interact/pivot-table-dialog.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const text = (value: string): CellValue => ({ kind: 'text', value });
const num = (value: number): CellValue => ({ kind: 'number', value });

const makeWb = () => {
  const calls: string[] = [];
  const wb = {
    capabilities: { pivotTableMutate: true },
    getPivotTables: () => [],
    getValue: ({ row, col }: { row: number; col: number }) => {
      if (row === 0 && col === 0) return text('Region');
      if (row === 0 && col === 1) return text('Sales');
      if (row === 1 && col === 0) return text('East');
      if (row === 1 && col === 1) return num(10);
      if (row === 2 && col === 0) return text('West');
      if (row === 2 && col === 1) return num(20);
      return { kind: 'blank' } as CellValue;
    },
    createPivotCache: () => 9,
    removePivotCache: () => true,
    addPivotCacheField: (_cache: number, name: string) => {
      calls.push(`field:${name}`);
      return calls.filter((c) => c.startsWith('field:')).length - 1;
    },
    addPivotCacheSharedItem: () => true,
    addPivotCacheRecord: () => 0,
    setPivotCacheRecordValue: () => true,
    createPivotTable: () => {
      calls.push('pivot');
      return 0;
    },
    removePivotTable: () => true,
    setPivotTableGrandTotals: (_sheet: number, _pivot: number, rows: boolean, cols: boolean) => {
      calls.push(`grand:${rows}:${cols}`);
      return true;
    },
    addPivotField: () => 0,
    setPivotFieldSort: () => true,
    addPivotDataField: () => 0,
  } as unknown as WorkbookHandle;
  return { wb, calls };
};

describe('attachPivotTableDialog', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  it('creates a PivotTable from the selected range', () => {
    const { wb, calls } = makeWb();
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 });
    const handle = attachPivotTableDialog({ host, store, wb, strings: en });

    handle.open();
    expect(document.body.textContent).toContain('Create PivotTable');
    expect(document.body.textContent).toContain('A1:B3');
    expect(document.querySelectorAll('.fc-pivotdlg__section')).toHaveLength(3);
    expect(document.querySelectorAll('.fc-pivotdlg__checkgrid .fc-pivotdlg__check')).toHaveLength(
      4,
    );

    const form = document.querySelector('form');
    expect(form).toBeTruthy();
    if (!form) throw new Error('missing PivotTable form');
    form.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));

    expect(calls).toContain('field:Region');
    expect(calls).toContain('field:Sales');
    expect(calls).toContain('pivot');
    expect(calls).toContain('grand:true:true');
    expect(document.querySelector('.fc-pivotdlg')?.hasAttribute('hidden')).toBe(true);
    handle.detach();
  });

  it('treats Enter as the default OK action when creation is available', () => {
    const { wb, calls } = makeWb();
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 });
    const handle = attachPivotTableDialog({ host, store, wb, strings: en });

    handle.open();
    const overlay = document.querySelector<HTMLElement>('.fc-pivotdlg');
    overlay?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));

    expect(calls).toContain('pivot');
    expect(overlay?.hasAttribute('hidden')).toBe(true);
    handle.detach();
  });

  it('shows a disabled-state message when mutation is unavailable', () => {
    const { wb } = makeWb();
    Object.defineProperty(wb, 'capabilities', { value: { pivotTableMutate: false } });
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 });
    const handle = attachPivotTableDialog({ host, store, wb, strings: en });

    handle.open();
    expect(document.body.textContent).toContain('does not support PivotTable creation');
    expect(document.querySelector('.fc-fmtdlg__btn--primary')?.hasAttribute('disabled')).toBe(true);
    expect(document.activeElement?.textContent).toBe('Cancel');
    handle.detach();
  });
});
