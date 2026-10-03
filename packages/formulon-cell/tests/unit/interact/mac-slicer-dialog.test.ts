import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings/en.js';
import { ja } from '../../../src/i18n/strings/ja.js';
import { attachMacSlicerDialog } from '../../../src/interact/mac-slicer-dialog.js';
import { createSpreadsheetStore, type SlicerSpec } from '../../../src/store/store.js';

const workbookWithTable = (): WorkbookHandle =>
  ({
    getTables: () => [
      {
        name: 'Sales',
        displayName: 'Sales',
        ref: 'A1:B3',
        sheetIndex: 0,
        columns: ['Region', 'Amount'],
      },
    ],
  }) as unknown as WorkbookHandle;

describe('attachMacSlicerDialog', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('offers workbook table columns and delegates the selected field', () => {
    const selected: Array<{ tableName: string; column: string }> = [];
    const handle = attachMacSlicerDialog({
      host,
      store: createSpreadsheetStore(),
      getWb: workbookWithTable,
      onAdd: (input) => {
        selected.push(input);
        return { id: 'slicer-1', ...input, selected: [] } satisfies SlicerSpec;
      },
    });
    handle.open();
    const column = host.ownerDocument.querySelector<HTMLSelectElement>('.fc-macslicerdlg__column');
    const ok = host.ownerDocument.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    expect(column).not.toBeNull();
    expect(ok).not.toBeNull();
    expect(column?.options.length).toBe(2);
    if (!column || !ok) throw new Error('slicer chooser controls missing');
    column.value = 'Amount';
    ok.click();
    expect(selected).toEqual([{ tableName: 'Sales', column: 'Amount' }]);
    expect(document.querySelector<HTMLElement>('.fc-macslicerdlg')?.hidden).toBe(true);
    handle.detach();
  });

  it('keeps the chooser unavailable when the workbook has no tables', () => {
    const handle = attachMacSlicerDialog({
      host,
      store: createSpreadsheetStore(),
      getWb: () => ({ getTables: () => [] }) as unknown as WorkbookHandle,
      strings: en,
      onAdd: () => null,
    });
    handle.open();
    const ok = host.ownerDocument.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    expect(ok?.disabled).toBe(true);
    expect(document.querySelector('[data-mac-slicer-error]')?.textContent).toContain('No table');
    handle.detach();
  });

  it('keeps the strings set last across refresh and relabels every control', () => {
    const handle = attachMacSlicerDialog({
      host,
      store: createSpreadsheetStore(),
      getWb: workbookWithTable,
      strings: ja,
      onAdd: () => null,
    });
    handle.setStrings(en);
    handle.refresh();
    const overlay = document.querySelector<HTMLElement>('.fc-macslicerdlg');
    expect(overlay?.getAttribute('aria-label')).toBe(en.slicer.addSlicer);
    expect(overlay?.querySelector('.fc-fmtdlg__header')?.textContent).toBe(en.slicer.addSlicer);
    expect(overlay?.querySelector('.fc-macslicerdlg__table')?.getAttribute('aria-label')).toBe(
      en.macSlicer.table,
    );
    expect(overlay?.querySelector('.fc-fmtdlg__btn--primary')?.textContent).toBe(en.macSlicer.ok);
    handle.detach();
  });

  it('applies the shared dialog class to the overlay without leaking it onto the panel', () => {
    const handle = attachMacSlicerDialog({
      host,
      store: createSpreadsheetStore(),
      getWb: workbookWithTable,
      onAdd: () => null,
    });
    const overlay = document.querySelector<HTMLElement>('.fc-macslicerdlg');
    expect(overlay?.classList.contains('fc-fmtdlg')).toBe(true);
    expect(
      overlay?.querySelector('.fc-macslicerdlg__panel')?.classList.contains('fc-macslicerdlg'),
    ).toBe(false);
    handle.detach();
  });
});
