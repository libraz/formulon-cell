import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { attachNamedRangeDialog } from '../../../src/interact/named-range-dialog.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

describe('attachNamedRangeDialog', () => {
  let host: HTMLElement;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    wb = await newWb();
  });

  afterEach(() => {
    wb.dispose();
    document.body.innerHTML = '';
  });

  it('mounts hidden, opens shows the empty-state', () => {
    const handle = attachNamedRangeDialog({ host, wb });
    const overlay = document.querySelector<HTMLElement>('.fc-namedlg');
    expect(overlay?.hidden).toBe(true);

    handle.open();
    expect(overlay?.hidden).toBe(false);
    const empty = document.querySelector<HTMLElement>('.fc-namedlg__empty');
    expect(empty?.textContent).toBeTruthy();
    handle.detach();
  });

  it('Escape closes the overlay', () => {
    const handle = attachNamedRangeDialog({ host, wb });
    handle.open();
    const overlay = document.querySelector<HTMLElement>('.fc-namedlg') as HTMLElement;
    overlay.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(overlay.hidden).toBe(true);
    handle.detach();
  });

  it('bindWorkbook re-renders the listing on next open', () => {
    const handle = attachNamedRangeDialog({ host, wb });
    handle.open();
    handle.close();
    // Swap to a different workbook handle. The stub engine never has defined
    // names, so we just confirm bindWorkbook + open don't throw.
    handle.bindWorkbook(wb);
    handle.open();
    expect(document.querySelector<HTMLElement>('.fc-namedlg__empty')?.textContent).toBeTruthy();
    handle.detach();
  });

  it('hides the add-form when capability is off (read-only fallback note shown)', () => {
    const handle = attachNamedRangeDialog({ host, wb });
    handle.open();
    expect(document.querySelector<HTMLElement>('.fc-namedlg__form')).toBeNull();
    const note = document.querySelector<HTMLElement>('.fc-namedlg__note');
    expect(note?.textContent).toBeTruthy();
    expect(note?.id).toBe('fc-namedlg-readonly-note');
    const [newBtn, editBtn, deleteBtn] = Array.from(
      document.querySelectorAll<HTMLButtonElement>('.fc-namedlg__actions button'),
    );
    const quickInput = document.querySelector<HTMLInputElement>('.fc-namedlg__refers-input');
    const quickCommit = document.querySelector<HTMLButtonElement>('.fc-namedlg__refers-icon');
    expect(quickCommit?.textContent).toBe('');
    expect(quickCommit?.classList.contains('fc-namedlg__refers-icon--commit')).toBe(true);
    for (const control of [newBtn, editBtn, deleteBtn, quickInput, quickCommit]) {
      expect(control?.disabled).toBe(true);
      expect(control?.getAttribute('aria-describedby')).toBe('fc-namedlg-readonly-note');
      expect(control?.title).toBe(note?.textContent);
    }
    handle.detach();
  });
});

interface MutableWb {
  capabilities: { definedNameMutate: boolean; definedNameScopes: boolean };
  readonly sheetCount: number;
  definedNames(): IterableIterator<{ name: string; formula: string; localSheetId: number }>;
  sheetName(index: number): string;
  getValue(addr: { sheet: number; row: number; col: number }): FakeCellValue;
  setDefinedNameEntry(name: string, formula: string, localSheetId?: number): boolean;
  recalc(): void;
}

type FakeCellValue =
  | { kind: 'blank' }
  | { kind: 'number'; value: number }
  | { kind: 'text'; value: string };

const makeMutableWb = (): {
  wb: WorkbookHandle;
  calls: { name: string; formula: string; localSheetId: number }[];
  registry: Map<string, string>;
  values: Map<string, FakeCellValue>;
} => {
  const calls: { name: string; formula: string; localSheetId: number }[] = [];
  const registry = new Map<string, string>();
  const values = new Map<string, FakeCellValue>();
  const fake: MutableWb = {
    capabilities: { definedNameMutate: true, definedNameScopes: true },
    get sheetCount() {
      return 1;
    },
    *definedNames() {
      for (const [name, formula] of registry) yield { name, formula, localSheetId: -1 };
    },
    sheetName(index: number): string {
      return index === 0 ? 'Sheet1' : `Sheet${index + 1}`;
    },
    getValue(addr) {
      return values.get(`${addr.sheet}:${addr.row}:${addr.col}`) ?? { kind: 'blank' };
    },
    setDefinedNameEntry(name: string, formula: string, localSheetId = -1): boolean {
      calls.push({ name, formula, localSheetId });
      if (formula === '') registry.delete(name);
      else registry.set(name, formula);
      return true;
    },
    recalc() {},
  };
  return { wb: fake as unknown as WorkbookHandle, calls, registry, values };
};

describe('attachNamedRangeDialog (mutate enabled)', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('projects disabled reasons when no defined name is selected', () => {
    const { wb } = makeMutableWb();
    const handle = attachNamedRangeDialog({ host, wb, strings: en });
    handle.open();
    const [newBtn, editBtn, deleteBtn, filterBtn] = Array.from(
      document.querySelectorAll<HTMLButtonElement>('.fc-namedlg__actions button'),
    );
    const quickInput = document.querySelector<HTMLInputElement>('.fc-namedlg__refers-input');
    const quickCommit = document.querySelector<HTMLButtonElement>('.fc-namedlg__refers-icon');
    const quickCancel = Array.from(
      document.querySelectorAll<HTMLButtonElement>('.fc-namedlg__refers-icon'),
    )[1];

    expect(quickCommit?.textContent).toBe('');
    expect(quickCancel?.textContent).toBe('');
    expect(quickCommit?.classList.contains('fc-namedlg__refers-icon--commit')).toBe(true);
    expect(quickCancel?.classList.contains('fc-namedlg__refers-icon--cancel')).toBe(true);
    expect(newBtn?.disabled).toBe(false);
    expect(editBtn?.disabled).toBe(true);
    expect(deleteBtn?.disabled).toBe(true);
    expect(filterBtn?.disabled).toBe(true);
    expect(editBtn?.dataset.disabledReason).toBe(en.namedRangeDialog.selectNameActionReason);
    expect(deleteBtn?.getAttribute('aria-description')).toBe(
      en.namedRangeDialog.selectNameActionReason,
    );
    expect(filterBtn?.dataset.disabledReason).toBe(en.namedRangeDialog.filterRequiresNames);
    expect(filterBtn?.title).toBe(
      `${en.namedRangeDialog.filterButton}\n${en.namedRangeDialog.filterRequiresNames}`,
    );
    expect(quickInput?.disabled).toBe(true);
    expect(quickCommit?.dataset.disabledReason).toBe(
      en.namedRangeDialog.quickRefersRequiresSelection,
    );
    expect(quickCancel?.dataset.disabledReason).toBe(
      en.namedRangeDialog.quickRefersRequiresSelection,
    );
    handle.detach();
  });

  it('New opens a New Name dialog and writes through on submit', () => {
    const { wb, calls } = makeMutableWb();
    const handle = attachNamedRangeDialog({ host, wb });
    handle.open();
    const newButton = document.querySelector<HTMLButtonElement>('.fc-namedlg__actions button');
    newButton?.click();
    const editor = document.querySelector<HTMLElement>('.fc-namedlg-editor');
    expect(editor?.hidden).toBe(false);
    expect(editor?.getAttribute('aria-label')).toBe('新しい名前');
    const form = document.querySelector<HTMLFormElement>('.fc-namedlg-editor__form');
    expect(form).not.toBeNull();
    const nameField = document.querySelector<HTMLInputElement>(
      '.fc-namedlg-editor__input[aria-label="名前"]',
    );
    const formulaField = document.querySelector<HTMLInputElement>(
      '.fc-namedlg-editor__input[aria-label="参照先"]',
    );
    expect(nameField?.getAttribute('aria-label')).toBeTruthy();
    expect(formulaField?.getAttribute('aria-label')).toBeTruthy();
    const scopeField = document.querySelector<HTMLSelectElement>(
      '.fc-namedlg-editor__input[aria-label="範囲"]',
    );
    expect(scopeField?.value).toBe('-1');
    expect(
      Array.from(scopeField?.options ?? [], (option) => [option.value, option.textContent]),
    ).toEqual([
      ['-1', 'ブック'],
      ['0', 'Sheet1'],
    ]);
    if (nameField) nameField.value = 'TaxRate';
    if (formulaField) formulaField.value = '=Sheet1!$A$1';
    form?.requestSubmit();
    expect(calls).toEqual([{ name: 'TaxRate', formula: '=Sheet1!$A$1', localSheetId: -1 }]);
    // Listing re-renders to reflect the new entry.
    const items = document.querySelectorAll('.fc-namedlg__item');
    expect(items.length).toBe(1);
    handle.detach();
  });

  it('passes the selected sheet scope when creating a defined name', () => {
    const { wb, calls } = makeMutableWb();
    const handle = attachNamedRangeDialog({ host, wb });
    handle.open();
    document.querySelector<HTMLButtonElement>('.fc-namedlg__actions button')?.click();
    const nameField = document.querySelector<HTMLInputElement>(
      '.fc-namedlg-editor__input[aria-label="名前"]',
    );
    const scopeField = document.querySelector<HTMLSelectElement>(
      '.fc-namedlg-editor__input[aria-label="範囲"]',
    );
    const formulaField = document.querySelector<HTMLInputElement>(
      '.fc-namedlg-editor__input[aria-label="参照先"]',
    );

    if (nameField) nameField.value = 'LocalRate';
    if (scopeField) scopeField.value = '0';
    if (formulaField) formulaField.value = '=Sheet1!$A$1';
    document.querySelector<HTMLFormElement>('.fc-namedlg-editor__form')?.requestSubmit();

    expect(calls).toEqual([{ name: 'LocalRate', formula: '=Sheet1!$A$1', localSheetId: 0 }]);
    handle.detach();
  });

  it('notifies the host after defined-name mutations so recalculated cells can re-project', () => {
    const { wb, registry } = makeMutableWb();
    registry.set('TaxRate', '=Sheet1!$A$1');
    const afterMutate = vi.fn();
    const handle = attachNamedRangeDialog({ host, wb, onAfterMutate: afterMutate });
    handle.open();

    document.querySelector<HTMLButtonElement>('.fc-namedlg__actions button')?.click();
    const nameField = document.querySelector<HTMLInputElement>(
      '.fc-namedlg-editor__input[aria-label="名前"]',
    );
    const formulaField = document.querySelector<HTMLInputElement>(
      '.fc-namedlg-editor__input[aria-label="参照先"]',
    );
    if (nameField) nameField.value = 'Discount';
    if (formulaField) formulaField.value = '=Sheet1!$B$1';
    document.querySelector<HTMLFormElement>('.fc-namedlg-editor__form')?.requestSubmit();

    expect(afterMutate).toHaveBeenCalledTimes(1);

    const firstRow = document.querySelector<HTMLButtonElement>('.fc-namedlg__item');
    firstRow?.click();
    const quickInput = document.querySelector<HTMLInputElement>('.fc-namedlg__refers-input');
    if (quickInput) quickInput.value = '=Sheet1!$C$3';
    document.querySelector<HTMLButtonElement>('.fc-namedlg__refers-icon')?.click();

    expect(afterMutate).toHaveBeenCalledTimes(2);
    handle.detach();
  });

  it('openNew opens the New Name editor directly', async () => {
    const { wb } = makeMutableWb();
    const handle = attachNamedRangeDialog({ host, wb });
    handle.openNew();
    await new Promise((resolve) => requestAnimationFrame(resolve));

    const manager = document.querySelector<HTMLElement>('.fc-namedlg');
    const editor = document.querySelector<HTMLElement>('.fc-namedlg-editor');
    expect(manager?.hidden).toBe(false);
    expect(editor?.hidden).toBe(false);
    expect(editor?.getAttribute('aria-label')).toBe('新しい名前');
    handle.detach();
  });

  it('records Name Manager edits in unified history', () => {
    const { wb, registry } = makeMutableWb();
    const history = new History();
    const handle = attachNamedRangeDialog({ host, wb, history });
    handle.open();
    document.querySelector<HTMLButtonElement>('.fc-namedlg__actions button')?.click();
    const nameField = document.querySelector<HTMLInputElement>(
      '.fc-namedlg-editor__input[aria-label="名前"]',
    );
    const formulaField = document.querySelector<HTMLInputElement>(
      '.fc-namedlg-editor__input[aria-label="参照先"]',
    );
    if (nameField) nameField.value = 'TaxRate';
    if (formulaField) formulaField.value = '=Sheet1!$A$1';
    document.querySelector<HTMLFormElement>('.fc-namedlg-editor__form')?.requestSubmit();

    expect(registry.get('TaxRate')).toBe('=Sheet1!$A$1');
    expect(history.canUndo()).toBe(true);

    history.undo();
    expect(registry.has('TaxRate')).toBe(false);

    history.redo();
    expect(registry.get('TaxRate')).toBe('=Sheet1!$A$1');
    handle.detach();
  });

  it('renders Excel-like Name Manager actions and table columns', () => {
    const { wb, registry } = makeMutableWb();
    registry.set('TaxRate', '=Sheet1!$A$1');
    const handle = attachNamedRangeDialog({ host, wb });
    handle.open();

    const actions = Array.from(
      document.querySelectorAll<HTMLButtonElement>('.fc-namedlg__actions button'),
    );
    expect(actions.map((button) => button.textContent)).toEqual([
      '新規...',
      '編集...',
      '削除',
      'フィルター',
    ]);
    const headers = Array.from(
      document.querySelectorAll<HTMLElement>('.fc-namedlg__head span'),
    ).map((cell) => cell.textContent);
    expect(headers).toEqual(['名前', '値', '参照先', '範囲', 'コメント']);
    const rowCells = Array.from(
      document.querySelectorAll<HTMLElement>('.fc-namedlg__item span'),
    ).map((cell) => cell.textContent);
    expect(rowCells).toEqual(['TaxRate', '-', '=Sheet1!$A$1', 'ブック', '']);
    handle.detach();
  });

  it('renders simple defined-name reference values in the Value column', () => {
    const { wb, registry, values } = makeMutableWb();
    registry.set('TaxRate', '=Sheet1!$A$1');
    values.set('0:0:0', { kind: 'number', value: 0.08 });
    const handle = attachNamedRangeDialog({ host, wb });
    handle.open();

    expect(document.querySelector<HTMLElement>('.fc-namedlg__row-value')?.textContent).toBe('0.08');
    handle.detach();
  });

  it('sorts the Name Manager list from column headers', () => {
    const { wb, registry } = makeMutableWb();
    registry.set('TaxRate', '=Sheet1!$C$1');
    registry.set('Budget', '=Sheet1!$A$1');
    const handle = attachNamedRangeDialog({ host, wb });
    handle.open();

    const rowNames = (): (string | null)[] =>
      Array.from(document.querySelectorAll<HTMLElement>('.fc-namedlg__row-name')).map(
        (cell) => cell.textContent,
      );
    expect(rowNames()).toEqual(['Budget', 'TaxRate']);
    const [nameHeader, , refersToHeader] = Array.from(
      document.querySelectorAll<HTMLButtonElement>('.fc-namedlg__head button'),
    );
    expect(nameHeader?.getAttribute('aria-sort')).toBe('ascending');

    nameHeader?.click();
    expect(rowNames()).toEqual(['TaxRate', 'Budget']);
    expect(nameHeader?.getAttribute('aria-sort')).toBe('descending');

    refersToHeader?.click();
    expect(rowNames()).toEqual(['Budget', 'TaxRate']);
    expect(refersToHeader?.getAttribute('aria-sort')).toBe('ascending');
    handle.detach();
  });

  it('Edit loads the selected name into the inline form for replacement', () => {
    const { wb, calls, registry } = makeMutableWb();
    registry.set('TaxRate', '=Sheet1!$A$1');
    const handle = attachNamedRangeDialog({ host, wb });
    handle.open();
    const edit = Array.from(
      document.querySelectorAll<HTMLButtonElement>('.fc-namedlg__actions button'),
    )[1];
    edit?.click();
    const editor = document.querySelector<HTMLElement>('.fc-namedlg-editor');
    expect(editor?.hidden).toBe(false);
    expect(editor?.getAttribute('aria-label')).toBe('名前の編集');
    const nameField = document.querySelector<HTMLInputElement>(
      '.fc-namedlg-editor__input[aria-label="名前"]',
    );
    const formulaField = document.querySelector<HTMLInputElement>(
      '.fc-namedlg-editor__input[aria-label="参照先"]',
    );
    expect(nameField?.value).toBe('TaxRate');
    expect(formulaField?.value).toBe('=Sheet1!$A$1');
    if (formulaField) formulaField.value = '=Sheet1!$B$2';
    document.querySelector<HTMLFormElement>('.fc-namedlg-editor__form')?.requestSubmit();
    expect(calls).toEqual([{ name: 'TaxRate', formula: '=Sheet1!$B$2', localSheetId: -1 }]);
    expect(editor?.hidden).toBe(true);
    handle.detach();
  });

  it('quick Refers to box commits and cancels selected-name reference edits', () => {
    const { wb, calls, registry } = makeMutableWb();
    registry.set('TaxRate', '=Sheet1!$A$1');
    const handle = attachNamedRangeDialog({ host, wb });
    handle.open();

    const quick = document.querySelector<HTMLInputElement>('.fc-namedlg__refers-input');
    const [commit, cancel] = Array.from(
      document.querySelectorAll<HTMLButtonElement>('.fc-namedlg__refers-icon'),
    );
    expect(quick?.value).toBe('=Sheet1!$A$1');
    if (quick) quick.value = '=Sheet1!$C$3';
    commit?.click();
    expect(calls).toEqual([{ name: 'TaxRate', formula: '=Sheet1!$C$3', localSheetId: -1 }]);
    expect(quick?.value).toBe('=Sheet1!$C$3');

    if (quick) quick.value = '=Sheet1!$D$4';
    cancel?.click();
    expect(calls).toEqual([{ name: 'TaxRate', formula: '=Sheet1!$C$3', localSheetId: -1 }]);
    expect(quick?.value).toBe('=Sheet1!$C$3');
    handle.detach();
  });

  it('uses shared range pickers for Name Manager reference inputs', () => {
    const { wb, registry } = makeMutableWb();
    registry.set('TaxRate', '=A1');
    let picked = '=B2:C4';
    const listeners: Array<() => void> = [];
    const handle = attachNamedRangeDialog({
      host,
      wb,
      getSelectedRangeFormula: () => picked,
      subscribeToRangeChanges: (listener) => {
        listeners.push(listener);
        return () => undefined;
      },
    });
    handle.open();

    const quickButton = document.querySelector<HTMLButtonElement>(
      '[data-range-picker="named-range-quick-refers-to"]',
    );
    const quickInput = document.querySelector<HTMLInputElement>('.fc-namedlg__refers-input');
    expect(quickButton?.getAttribute('aria-label')).toBe('範囲の選択');
    quickButton?.click();
    expect(quickInput?.value).toBe('=B2:C4');
    expect(quickButton?.getAttribute('aria-pressed')).toBe('true');
    expect(document.querySelector('.fc-namedlg.fc-fmtdlg--range-picking')).toBeTruthy();
    picked = '=D5:D8';
    listeners.at(-1)?.();
    expect(quickInput?.value).toBe('=D5:D8');

    document.querySelector<HTMLButtonElement>('.fc-namedlg__actions button')?.click();
    const editorButton = document.querySelector<HTMLButtonElement>(
      '[data-range-picker="named-range-editor-refers-to"]',
    );
    const editorInput = document.querySelector<HTMLInputElement>(
      '.fc-namedlg-editor__input[aria-label="参照先"]',
    );
    editorButton?.click();
    expect(quickButton?.getAttribute('aria-pressed')).toBe('false');
    expect(editorButton?.getAttribute('aria-pressed')).toBe('true');
    picked = '=E1:F2';
    listeners.at(-1)?.();
    expect(editorInput?.value).toBe('=E1:F2');

    handle.detach();
  });

  it('Filter menu narrows the Name Manager list to names with formula errors', () => {
    const { wb, registry } = makeMutableWb();
    registry.set('TaxRate', '=Sheet1!$A$1');
    registry.set('BrokenRef', '=#REF!');
    const handle = attachNamedRangeDialog({ host, wb });
    handle.open();

    const filter = Array.from(
      document.querySelectorAll<HTMLButtonElement>('.fc-namedlg__actions button'),
    )[3];
    filter?.click();
    const items = Array.from(
      document.querySelectorAll<HTMLButtonElement>('.fc-namedlg__filter-menu [role="menuitem"]'),
    );
    expect(items.map((item) => item.textContent)).toEqual([
      'フィルターのクリア',
      'エラーのある名前',
      'エラーのない名前',
      'ブック範囲の名前',
    ]);
    items[1]?.click();

    const rows = Array.from(document.querySelectorAll<HTMLElement>('.fc-namedlg__item'));
    expect(rows).toHaveLength(1);
    expect(rows[0]?.textContent).toContain('BrokenRef');
    expect(filter?.textContent).toBe('フィルター: エラーのある名前');

    filter?.click();
    document
      .querySelectorAll<HTMLButtonElement>('.fc-namedlg__filter-menu [role="menuitem"]')[0]
      ?.click();
    expect(document.querySelectorAll<HTMLElement>('.fc-namedlg__item')).toHaveLength(2);
    handle.detach();
  });

  it('rejects empty name with inline error and no engine call', () => {
    const { wb, calls } = makeMutableWb();
    const handle = attachNamedRangeDialog({ host, wb });
    handle.open();
    document.querySelector<HTMLButtonElement>('.fc-namedlg__actions button')?.click();
    document.querySelector<HTMLFormElement>('.fc-namedlg-editor__form')?.requestSubmit();
    expect(calls).toEqual([]);
    expect(
      document.querySelector<HTMLElement>('.fc-namedlg-editor .fc-namedlg__error')?.hidden,
    ).toBe(false);
    handle.detach();
  });

  it('rejects empty reference with inline error and no engine call', () => {
    const { wb, calls } = makeMutableWb();
    const handle = attachNamedRangeDialog({ host, wb });
    handle.open();
    document.querySelector<HTMLButtonElement>('.fc-namedlg__actions button')?.click();
    const nameField = document.querySelector<HTMLInputElement>(
      '.fc-namedlg-editor__input[aria-label="名前"]',
    );
    if (nameField) nameField.value = 'TaxRate';
    document.querySelector<HTMLFormElement>('.fc-namedlg-editor__form')?.requestSubmit();
    expect(calls).toEqual([]);
    expect(
      document.querySelector<HTMLElement>('.fc-namedlg-editor .fc-namedlg__error')?.hidden,
    ).toBe(false);
    handle.detach();
  });

  it('Delete button passes empty formula (engine convention) and refreshes', () => {
    const { wb, calls, registry } = makeMutableWb();
    registry.set('TaxRate', '=Sheet1!$A$1');
    const handle = attachNamedRangeDialog({ host, wb });
    handle.open();
    const del = Array.from(
      document.querySelectorAll<HTMLButtonElement>('.fc-namedlg__actions button'),
    )[2];
    expect(del).not.toBeNull();
    del?.click();
    const confirm = document.querySelector<HTMLElement>('.fc-namedlg-confirm');
    expect(confirm?.hidden).toBe(false);
    expect(confirm?.textContent).toContain('TaxRate');
    expect(calls).toEqual([]);
    const ok = Array.from(
      document.querySelectorAll<HTMLButtonElement>('.fc-namedlg-confirm button'),
    ).find((button) => button.textContent === 'OK');
    ok?.click();
    expect(calls).toEqual([{ name: 'TaxRate', formula: '', localSheetId: -1 }]);
    // List should now be empty.
    expect(document.querySelector<HTMLElement>('.fc-namedlg__empty')?.textContent).toBeTruthy();
    handle.detach();
  });

  it('Cancel in delete confirmation leaves the selected name intact', () => {
    const { wb, calls, registry } = makeMutableWb();
    registry.set('TaxRate', '=Sheet1!$A$1');
    const handle = attachNamedRangeDialog({ host, wb });
    handle.open();
    Array.from(
      document.querySelectorAll<HTMLButtonElement>('.fc-namedlg__actions button'),
    )[2]?.click();
    const cancel = Array.from(
      document.querySelectorAll<HTMLButtonElement>('.fc-namedlg-confirm button'),
    ).find((button) => button.textContent === 'キャンセル');
    cancel?.click();

    expect(calls).toEqual([]);
    expect(document.querySelectorAll<HTMLElement>('.fc-namedlg__item')).toHaveLength(1);
    expect(document.querySelector<HTMLElement>('.fc-namedlg-confirm')?.hidden).toBe(true);
    handle.detach();
  });

  it('exposes a selectable listbox and moves row selection with keyboard', () => {
    const { wb, registry } = makeMutableWb();
    registry.set('TaxRate', '=Sheet1!$A$1');
    registry.set('Budget', '=Sheet1!$B$1');
    const handle = attachNamedRangeDialog({ host, wb });
    handle.open();

    const list = document.querySelector<HTMLElement>('.fc-namedlg__list');
    expect(list?.getAttribute('role')).toBe('listbox');
    const rows = Array.from(document.querySelectorAll<HTMLElement>('.fc-namedlg__item'));
    expect(rows).toHaveLength(2);
    expect(rows[0]?.getAttribute('role')).toBe('option');
    expect(rows[0]?.getAttribute('aria-selected')).toBe('true');
    expect(rows[0]?.tabIndex).toBe(0);
    expect(rows[1]?.getAttribute('aria-selected')).toBe('false');
    expect(rows[1]?.tabIndex).toBe(-1);

    rows[0]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    expect(rows[1]?.getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement).toBe(rows[1]);

    rows[1]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true }));
    expect(rows[0]?.getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement).toBe(rows[0]);

    rows[0]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
    expect(rows[1]?.getAttribute('aria-selected')).toBe('true');
    expect(document.activeElement).toBe(rows[1]);
    handle.detach();
  });

  it('Delete key removes the selected defined-name row', () => {
    const { wb, calls, registry } = makeMutableWb();
    registry.set('TaxRate', '=Sheet1!$A$1');
    registry.set('Budget', '=Sheet1!$B$1');
    const handle = attachNamedRangeDialog({ host, wb });
    handle.open();

    const firstRow = document.querySelector<HTMLElement>('.fc-namedlg__item');
    firstRow?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', bubbles: true }));

    const ok = Array.from(
      document.querySelectorAll<HTMLButtonElement>('.fc-namedlg-confirm button'),
    ).find((button) => button.textContent === 'OK');
    ok?.click();

    expect(calls).toEqual([{ name: 'Budget', formula: '', localSheetId: -1 }]);
    const rows = Array.from(document.querySelectorAll<HTMLElement>('.fc-namedlg__item'));
    expect(rows).toHaveLength(1);
    expect(rows[0]?.textContent).toContain('TaxRate');
    expect(rows[0]?.getAttribute('aria-selected')).toBe('true');
    handle.detach();
  });

  it('keeps Name Manager sort and filter menu buttons on the shared dialog primitive', () => {
    const source = readFileSync(join(root, 'src/interact/named-range-dialog.ts'), 'utf8');
    expect(source).toContain('function createNameManagerButton(');
    expect(source).toContain("createNameManagerButton(label, 'fc-namedlg__sort')");
    expect(source).toContain(
      "createNameManagerButton(filterLabel(filter), 'fc-namedlg__filter-item'",
    );
    expect(source).not.toContain("document.createElement('button')");
  });

  it('keeps Name Manager list and filter menu close to Excel 365 desktop geometry', () => {
    const css = readFileSync(
      join(root, 'src/styles/core/app/dialog-modules/conditional-and-names.css'),
      'utf8',
    );

    expect(css).toMatch(
      /\.fc-namedlg__filter-menu\s*\{[\s\S]*?min-width: 210px;[\s\S]*?border-radius: 2px;[\s\S]*?box-shadow:[\s\S]*?padding: 4px 0;/,
    );
    expect(css).toMatch(
      /\.fc-namedlg__filter-item\s*\{[\s\S]*?min-height: 24px;[\s\S]*?padding: 3px 26px 3px 24px;/,
    );
    expect(css).toMatch(
      /\.fc-namedlg__filter-item\[aria-checked="true"\]::before\s*\{[\s\S]*?border-bottom: 2px solid var\(--fc-accent, #107c41\);[\s\S]*?border-left: 2px solid var\(--fc-accent, #107c41\);[\s\S]*?content: "";[\s\S]*?transform: rotate\(-45deg\);/,
    );
    expect(css).toMatch(
      /\.fc-namedlg__sort--active::after\s*\{[\s\S]*?border: solid currentColor;[\s\S]*?content: "";/,
    );
    expect(css).toMatch(
      /\.fc-namedlg__sort\[data-sort-dir="asc"\]::after\s*\{[\s\S]*?rotate\(-45deg\);/,
    );
    expect(css).toMatch(
      /\.fc-namedlg__sort\[data-sort-dir="desc"\]::after\s*\{[\s\S]*?rotate\(135deg\);/,
    );
    expect(css).not.toContain('content: attr(data-sort-dir)');
    expect(css).toMatch(
      /\.fc-namedlg__refers-icon--commit::before\s*\{[\s\S]*?border-bottom: 2px solid currentColor;[\s\S]*?border-left: 2px solid currentColor;/,
    );
    expect(css).toMatch(
      /\.fc-namedlg__refers-icon::before,[\s\S]*?\.fc-namedlg__refers-icon::after\s*\{[\s\S]*?content: "";/,
    );
    expect(css).toMatch(
      /\.fc-namedlg__refers-icon--cancel::before,[\s\S]*?\.fc-namedlg__refers-icon--cancel::after\s*\{[\s\S]*?background: currentColor;/,
    );
    expect(css).not.toContain('content: "✓"');
    expect(css).toMatch(
      /\.fc-conddlg__item:hover,[\s\S]*?\.fc-namedlg__item:hover\s*\{[\s\S]*?background: var\(--fc-bg-hover/,
    );
    expect(css).toMatch(
      /\.fc-namedlg__item\[aria-selected="true"\],[\s\S]*?\.fc-namedlg__item--selected\s*\{[\s\S]*?background: var\(--fc-bg-hover/,
    );
    expect(css).not.toContain(
      '.fc-namedlg__item--selected {\n    background: var(--fc-accent-soft',
    );
  });

  it('keeps New/Edit Name child dialogs on compact desktop form geometry', () => {
    const css = readFileSync(
      join(root, 'src/styles/core/app/dialog-modules/conditional-and-names.css'),
      'utf8',
    );

    expect(css).toMatch(/\.fc-namedlg-editor__body\s*\{[\s\S]*?padding: 12px 14px 14px;/);
    expect(css).toMatch(/\.fc-namedlg-editor__form\s*\{[\s\S]*?gap: 6px;/);
    expect(css).toMatch(
      /\.fc-namedlg-editor__row\s*\{[\s\S]*?grid-template-columns: 72px minmax\(0, 1fr\);[\s\S]*?gap: 8px;/,
    );
    expect(css).toMatch(
      /\.fc-namedlg-editor__buttons\s*\{[\s\S]*?gap: 6px;[\s\S]*?margin-top: 6px;/,
    );
    expect(css).toMatch(/\.fc-namedlg-confirm__body\s*\{[\s\S]*?gap: 12px;[\s\S]*?padding: 14px;/);
    expect(css).toMatch(/\.fc-namedlg-confirm__buttons\s*\{[\s\S]*?gap: 6px;/);
  });
});
