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

  const fields = (): HTMLInputElement[] =>
    Array.from(
      host.ownerDocument.querySelectorAll<HTMLInputElement>(
        '.fc-macslicerdlg__field input[type="checkbox"]',
      ),
    );

  it('lists every column of the table as an unchecked checkbox, in source order', () => {
    const handle = attachMacSlicerDialog({
      host,
      store: createSpreadsheetStore(),
      getWb: workbookWithTable,
      strings: ja,
      onAdd: () => null,
    });
    handle.open();
    expect(
      Array.from(
        host.ownerDocument.querySelectorAll('.fc-macslicerdlg__field'),
        (row) => row.textContent,
      ),
    ).toEqual(['Region', 'Amount']);
    expect(fields().every((box) => !box.checked)).toBe(true);
    expect(host.ownerDocument.querySelector('.fc-fmtdlg__header')?.textContent).toBe(
      'スライサーの挿入',
    );
    expect(host.ownerDocument.querySelector('.fc-macslicerdlg__table')).toBeNull();
    handle.detach();
  });

  it('inserts one slicer per checked field', () => {
    const selected: Array<{ tableName: string; column: string }> = [];
    const handle = attachMacSlicerDialog({
      host,
      store: createSpreadsheetStore(),
      getWb: workbookWithTable,
      onAdd: (input) => {
        selected.push(input);
        return { id: `slicer-${selected.length}`, ...input, selected: [] } satisfies SlicerSpec;
      },
    });
    handle.open();
    for (const box of fields()) box.click();
    host.ownerDocument.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    expect(selected).toEqual([
      { tableName: 'Sales', column: 'Region' },
      { tableName: 'Sales', column: 'Amount' },
    ]);
    expect(document.querySelector<HTMLElement>('.fc-macslicerdlg')?.hidden).toBe(true);
    handle.detach();
  });

  it('closes without inserting when nothing is checked', () => {
    const selected: unknown[] = [];
    const handle = attachMacSlicerDialog({
      host,
      store: createSpreadsheetStore(),
      getWb: workbookWithTable,
      onAdd: (input) => {
        selected.push(input);
        return null;
      },
    });
    handle.open();
    host.ownerDocument.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    expect(selected).toEqual([]);
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
