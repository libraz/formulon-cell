import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import type { WriteResult } from '../../../src/engine/result.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings.js';
import { attachNamedRangeDialog } from '../../../src/interact/named-range-dialog.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const PRINT_AREA = '_xlnm.Print_Area';

interface StoredEntry {
  name: string;
  formula: string;
  localSheetId: number;
}

/**
 * Handle double implementing only the stored-key rule: a typed `print_area`
 * (any case) is stored as `_xlnm.Print_Area` at the scope of the sheet its
 * formula names; removal resolves the typed spelling the same way.
 */
const makeStoredKeyWb = (): {
  wb: WorkbookHandle;
  entries: Map<string, StoredEntry>;
  writes: { name: string; formula: string; localSheetId: number }[];
} => {
  const entries = new Map<string, StoredEntry>();
  const writes: { name: string; formula: string; localSheetId: number }[] = [];
  const sheets = ['Sheet1', 'Sheet2'];
  const resolve = (typed: string): string =>
    typed.toLowerCase() === 'print_area' ? PRINT_AREA : typed;
  const keyOf = (name: string, scope: number): string => `${scope}:${name}`;
  const fake = {
    capabilities: { definedNameMutate: true, definedNameScopes: true },
    sheetCount: sheets.length,
    *definedNames() {
      yield* entries.values();
    },
    sheetName: (i: number) => sheets[i] ?? '',
    sheetIndexByName: (n: string) => {
      const i = sheets.findIndex((s) => s.toLowerCase() === n.toLowerCase());
      return i < 0 ? null : i;
    },
    getValue: () => ({ kind: 'blank' }),
    resolveDefinedNameKey(name: string, formula: string, localSheetId = -1) {
      const stored = resolve(name);
      let scope = localSheetId;
      if (stored === PRINT_AREA) {
        const sheet = fake.sheetIndexByName(formula.replace(/^=/, '').split('!')[0] ?? '');
        if (sheet !== null) scope = sheet;
      }
      return { ok: true as const, name: stored, localSheetId: scope };
    },
    localizeDefinedName: (name: string) => name,
    setDefinedNameEntry(name: string, formula: string, localSheetId = -1): WriteResult {
      writes.push({ name, formula, localSheetId });
      const stored = resolve(name);
      if (formula === '') {
        entries.delete(keyOf(stored, localSheetId));
        return { ok: true };
      }
      let scope = localSheetId;
      if (stored === PRINT_AREA) {
        const sheet = fake.sheetIndexByName(formula.replace(/^=/, '').split('!')[0] ?? '');
        if (sheet !== null) scope = sheet;
      }
      entries.set(keyOf(stored, scope), { name: stored, formula, localSheetId: scope });
      return { ok: true };
    },
    recalc() {},
    recalcAuto() {},
  };
  return { wb: fake as unknown as WorkbookHandle, entries, writes };
};

describe('named range dialog with a stored-key spelling', () => {
  let host: HTMLElement;
  let handle: ReturnType<typeof attachNamedRangeDialog> | null = null;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
  });

  afterEach(() => {
    handle?.detach();
    handle = null;
    host.remove();
  });

  const editorFields = () => ({
    name: document.querySelector<HTMLInputElement>(
      '.fc-namedlg-editor__input[aria-label="名前"]',
    ) as HTMLInputElement,
    formula: document.querySelector<HTMLInputElement>(
      '.fc-namedlg-editor__input[aria-label="参照範囲"]',
    ) as HTMLInputElement,
    form: document.querySelector('.fc-namedlg-editor__form') as HTMLFormElement,
    error: document.querySelector('.fc-namedlg-editor .fc-namedlg__error') as HTMLElement,
  });

  it('keeps the entry a rename wrote when it resolves to the entry being replaced', () => {
    const { wb, entries } = makeStoredKeyWb();
    entries.set(`0:${PRINT_AREA}`, {
      name: PRINT_AREA,
      formula: '=Sheet1!$A$1:$B$2',
      localSheetId: 0,
    });
    handle = attachNamedRangeDialog({ store: createSpreadsheetStore(), strings: ja, host, wb });
    handle.open();
    document.querySelectorAll<HTMLButtonElement>('.fc-namedlg__actions button')[1]?.click();
    const f = editorFields();
    f.name.value = 'print_area';
    f.formula.value = '=Sheet1!$C$3:$D$4';
    f.form.requestSubmit();
    expect([...entries.values()]).toEqual([
      { name: PRINT_AREA, formula: '=Sheet1!$C$3:$D$4', localSheetId: 0 },
    ]);
  });

  it('refuses a name that resolves to an existing entry stored under another key', () => {
    const { wb, entries, writes } = makeStoredKeyWb();
    entries.set(`0:${PRINT_AREA}`, {
      name: PRINT_AREA,
      formula: '=Sheet1!$A$1:$B$2',
      localSheetId: 0,
    });
    handle = attachNamedRangeDialog({ store: createSpreadsheetStore(), strings: ja, host, wb });
    handle.open();
    document.querySelector<HTMLButtonElement>('.fc-namedlg__actions button')?.click();
    const f = editorFields();
    f.name.value = 'print_area';
    f.formula.value = '=Sheet1!$C$3:$D$4';
    f.form.requestSubmit();
    expect(writes).toEqual([]);
    expect(f.error.textContent).toBe(
      '入力した名前は既に存在します。一意の名前を入力してください。',
    );
  });
});
