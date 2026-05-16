import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachNamedRangeDialog } from '../../../src/interact/named-range-dialog.js';

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
    expect(document.querySelector<HTMLElement>('.fc-namedlg__note')?.textContent).toBeTruthy();
    handle.detach();
  });
});

interface MutableWb {
  capabilities: { definedNameMutate: boolean };
  definedNames(): IterableIterator<{ name: string; formula: string }>;
  setDefinedNameEntry(name: string, formula: string): boolean;
}

const makeMutableWb = (): {
  wb: WorkbookHandle;
  calls: { name: string; formula: string }[];
  registry: Map<string, string>;
} => {
  const calls: { name: string; formula: string }[] = [];
  const registry = new Map<string, string>();
  const fake: MutableWb = {
    capabilities: { definedNameMutate: true },
    *definedNames() {
      for (const [name, formula] of registry) yield { name, formula };
    },
    setDefinedNameEntry(name: string, formula: string): boolean {
      calls.push({ name, formula });
      if (formula === '') registry.delete(name);
      else registry.set(name, formula);
      return true;
    },
  };
  return { wb: fake as unknown as WorkbookHandle, calls, registry };
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
      '.fc-namedlg-editor__input[aria-label="参照"]',
    );
    expect(nameField?.getAttribute('aria-label')).toBeTruthy();
    expect(formulaField?.getAttribute('aria-label')).toBeTruthy();
    if (nameField) nameField.value = 'TaxRate';
    if (formulaField) formulaField.value = '=Sheet1!$A$1';
    form?.requestSubmit();
    expect(calls).toEqual([{ name: 'TaxRate', formula: '=Sheet1!$A$1' }]);
    // Listing re-renders to reflect the new entry.
    const items = document.querySelectorAll('.fc-namedlg__item');
    expect(items.length).toBe(1);
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
      '.fc-namedlg-editor__input[aria-label="参照"]',
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
      '新規作成...',
      '編集...',
      '削除',
      'フィルター',
    ]);
    const headers = Array.from(
      document.querySelectorAll<HTMLElement>('.fc-namedlg__head span'),
    ).map((cell) => cell.textContent);
    expect(headers).toEqual(['名前', '値', '参照', '範囲', 'コメント']);
    const rowCells = Array.from(
      document.querySelectorAll<HTMLElement>('.fc-namedlg__item span'),
    ).map((cell) => cell.textContent);
    expect(rowCells).toEqual(['TaxRate', '-', '=Sheet1!$A$1', 'ブック', '']);
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
      '.fc-namedlg-editor__input[aria-label="参照"]',
    );
    expect(nameField?.value).toBe('TaxRate');
    expect(formulaField?.value).toBe('=Sheet1!$A$1');
    if (formulaField) formulaField.value = '=Sheet1!$B$2';
    document.querySelector<HTMLFormElement>('.fc-namedlg-editor__form')?.requestSubmit();
    expect(calls).toEqual([{ name: 'TaxRate', formula: '=Sheet1!$B$2' }]);
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
    expect(calls).toEqual([{ name: 'TaxRate', formula: '=Sheet1!$C$3' }]);
    expect(quick?.value).toBe('=Sheet1!$C$3');

    if (quick) quick.value = '=Sheet1!$D$4';
    cancel?.click();
    expect(calls).toEqual([{ name: 'TaxRate', formula: '=Sheet1!$C$3' }]);
    expect(quick?.value).toBe('=Sheet1!$C$3');
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
    expect(calls).toEqual([{ name: 'TaxRate', formula: '' }]);
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

    expect(calls).toEqual([{ name: 'Budget', formula: '' }]);
    const rows = Array.from(document.querySelectorAll<HTMLElement>('.fc-namedlg__item'));
    expect(rows).toHaveLength(1);
    expect(rows[0]?.textContent).toContain('TaxRate');
    expect(rows[0]?.getAttribute('aria-selected')).toBe('true');
    handle.detach();
  });
});
