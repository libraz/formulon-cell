import { afterEach, beforeEach, describe, expect, it } from 'vitest';
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

  it('shows the add-form and writes through on submit', () => {
    const { wb, calls } = makeMutableWb();
    const handle = attachNamedRangeDialog({ host, wb });
    handle.open();
    const form = document.querySelector<HTMLFormElement>('.fc-namedlg__form');
    expect(form).not.toBeNull();
    const inputs = Array.from(document.querySelectorAll<HTMLInputElement>('.fc-namedlg__input'));
    expect(inputs.length).toBe(2);
    const [nameField, formulaField] = inputs;
    if (nameField) nameField.value = 'TaxRate';
    if (formulaField) formulaField.value = '=Sheet1!$A$1';
    form?.requestSubmit();
    expect(calls).toEqual([{ name: 'TaxRate', formula: '=Sheet1!$A$1' }]);
    // Listing re-renders to reflect the new entry.
    const items = document.querySelectorAll('.fc-namedlg__item');
    expect(items.length).toBe(1);
    handle.detach();
  });

  it('rejects empty name with inline error and no engine call', () => {
    const { wb, calls } = makeMutableWb();
    const handle = attachNamedRangeDialog({ host, wb });
    handle.open();
    const form = document.querySelector<HTMLFormElement>('.fc-namedlg__form');
    form?.requestSubmit();
    expect(calls).toEqual([]);
    expect(document.querySelector<HTMLElement>('.fc-namedlg__error')?.hidden).toBe(false);
    handle.detach();
  });

  it('rejects empty reference with inline error and no engine call', () => {
    const { wb, calls } = makeMutableWb();
    const handle = attachNamedRangeDialog({ host, wb });
    handle.open();
    const inputs = Array.from(document.querySelectorAll<HTMLInputElement>('.fc-namedlg__input'));
    const [nameField] = inputs;
    if (nameField) nameField.value = 'TaxRate';
    document.querySelector<HTMLFormElement>('.fc-namedlg__form')?.requestSubmit();
    expect(calls).toEqual([]);
    expect(document.querySelector<HTMLElement>('.fc-namedlg__error')?.hidden).toBe(false);
    handle.detach();
  });

  it('Delete button passes empty formula (engine convention) and refreshes', () => {
    const { wb, calls, registry } = makeMutableWb();
    registry.set('TaxRate', '=Sheet1!$A$1');
    const handle = attachNamedRangeDialog({ host, wb });
    handle.open();
    const del = document.querySelector<HTMLButtonElement>('.fc-namedlg__del');
    expect(del).not.toBeNull();
    del?.click();
    expect(calls).toEqual([{ name: 'TaxRate', formula: '' }]);
    // List should now be empty.
    expect(document.querySelector<HTMLElement>('.fc-namedlg__empty')?.textContent).toBeTruthy();
    handle.detach();
  });
});
