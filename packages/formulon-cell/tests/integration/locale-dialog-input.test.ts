import { afterEach, describe, expect, it } from 'vitest';
import { parseLocalizedNumericInput } from '../../src/commands/coerce-input.js';
import { addrKey } from '../../src/engine/address.js';
import { syncConditionalRulesToEngine } from '../../src/engine/cf-writeback.js';
import { hydrateValidationsFromEngine } from '../../src/engine/validation-sync.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { de, en, fr, ja, ko, th, zh } from '../../src/i18n/strings.js';
import { attachConditionalDialog } from '../../src/interact/conditional-dialog.js';
import { attachFormatDialog } from '../../src/interact/format-dialog.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../src/store/store.js';

const LOCALES = [
  ['en-US', 'windows-en_US', en],
  ['ja-JP', 'windows-ja_JP', ja],
  ['de-DE', 'windows-de_DE', de],
  ['fr-FR', 'windows-fr_FR', fr],
  ['zh-CN', 'windows-zh_CN', zh],
  ['ko-KR', 'windows-ko_KR', ko],
  ['th-TH', 'windows-th_TH', th],
] as const;

const setSelection = (store: SpreadsheetStore): void => {
  store.setState((state) => ({
    ...state,
    selection: {
      ...state.selection,
      active: { sheet: 0, row: 0, col: 0 },
      anchor: { sheet: 0, row: 0, col: 0 },
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
    },
  }));
};

const clickButton = (selector: string): void => {
  const button = document.querySelector<HTMLButtonElement>(selector);
  if (!button) throw new Error(`button not found: ${selector}`);
  button.click();
};

const inputWithLabel = (label: string): HTMLInputElement => {
  const row = [...document.querySelectorAll<HTMLLabelElement>('.fc-fmtdlg__row')].find(
    (candidate) => candidate.firstElementChild?.textContent?.trim() === label,
  );
  const input = row?.querySelector<HTMLInputElement>('input');
  if (!input) throw new Error(`input not found: ${label}`);
  return input;
};

afterEach(() => {
  document.body.innerHTML = '';
});

describe('locale-sensitive dialog boundaries', () => {
  it.each([1e-20, -1e-20, 1e20, 1.2345678901234567])(
    'preserves threshold precision when reopening %s',
    async (value) => {
      const store = createSpreadsheetStore();
      setSelection(store);
      store.setState((state) => ({
        ...state,
        conditional: {
          rules: [
            {
              kind: 'cell-value',
              op: '>',
              a: value,
              range: state.selection.range,
              apply: { fill: '#ff0000' },
            },
          ],
        },
      }));
      const host = document.createElement('div');
      document.body.append(host);
      const handle = attachConditionalDialog({
        host,
        store,
        strings: de,
        getLocale: () => 'de-DE',
      });
      try {
        handle.open({ mode: 'edit', editIndex: 0 });
        expect(
          parseLocalizedNumericInput(inputWithLabel(de.conditionalDialog.valueA).value, 'de-DE'),
        ).toBe(value);
      } finally {
        handle.detach();
      }
    },
  );

  it.each(LOCALES)('parses a cell-value threshold under %s', async (locale, profile, strings) => {
    const workbook = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    setSelection(store);
    const host = document.createElement('div');
    document.body.appendChild(host);
    try {
      expect(workbook.isStub).toBe(false);
      expect(workbook.setSpreadsheetProfileId(profile)).toBe(true);
      const handle = attachConditionalDialog({
        host,
        store,
        strings,
        getWb: () => workbook,
        getLocale: () => locale,
      });
      handle.open({ mode: 'new', kind: 'cell-value' });

      const valueA = inputWithLabel(strings.conditionalDialog.valueA);
      valueA.value = locale.startsWith('de') || locale.startsWith('fr') ? '1,5' : '1.5';
      clickButton('.fc-conddlg__addrow .fc-fmtdlg__btn--primary');

      expect(store.getState().conditional.rules[0]).toMatchObject({ kind: 'cell-value', a: 1.5 });
      expect(syncConditionalRulesToEngine(workbook, store.getState().conditional.rules, 0)).toEqual(
        {
          written: 1,
          skipped: 0,
        },
      );
      expect(workbook.getConditionalFormats(0)[0]?.formula1).toBe('1.5');
      handle.detach();
    } finally {
      workbook.dispose();
    }
  });

  it.each(LOCALES)('parses color-scale thresholds under %s', async (locale, profile, strings) => {
    const workbook = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    setSelection(store);
    const host = document.createElement('div');
    document.body.appendChild(host);
    try {
      expect(workbook.setSpreadsheetProfileId(profile)).toBe(true);
      const handle = attachConditionalDialog({
        host,
        store,
        strings,
        getWb: () => workbook,
        getLocale: () => locale,
      });
      handle.open({ mode: 'new', kind: 'color-scale' });
      const minType = document.querySelector<HTMLSelectElement>(
        `select[aria-label="${strings.conditionalDialog.stopMin} ${strings.conditionalDialog.scaleType}"]`,
      );
      const maxType = document.querySelector<HTMLSelectElement>(
        `select[aria-label="${strings.conditionalDialog.stopMax} ${strings.conditionalDialog.scaleType}"]`,
      );
      const minValue = document.querySelector<HTMLInputElement>(
        `input[aria-label="${strings.conditionalDialog.stopMin} ${strings.conditionalDialog.scaleValue}"]`,
      );
      const maxValue = document.querySelector<HTMLInputElement>(
        `input[aria-label="${strings.conditionalDialog.stopMax} ${strings.conditionalDialog.scaleValue}"]`,
      );
      if (!minType || !maxType || !minValue || !maxValue) {
        throw new Error('color-scale threshold controls missing');
      }
      minType.value = 'number';
      minType.dispatchEvent(new Event('change', { bubbles: true }));
      maxType.value = 'number';
      maxType.dispatchEvent(new Event('change', { bubbles: true }));
      const decimal = locale.startsWith('de') || locale.startsWith('fr') ? ',' : '.';
      minValue.value = `1${decimal}5`;
      maxValue.value = `2${decimal}5`;
      clickButton('.fc-conddlg__addrow .fc-fmtdlg__btn--primary');

      expect(store.getState().conditional.rules[0]).toMatchObject({
        kind: 'color-scale',
        thresholds: [
          { kind: 'number', value: 1.5 },
          { kind: 'number', value: 2.5 },
        ],
      });
      expect(syncConditionalRulesToEngine(workbook, store.getState().conditional.rules, 0)).toEqual(
        {
          written: 1,
          skipped: 0,
        },
      );
      expect(workbook.getConditionalFormats(0)[0]?.colorScale?.thresholds).toHaveLength(2);
      handle.detach();
    } finally {
      workbook.dispose();
    }
  });

  it.each([
    ['de-DE', 'windows-de_DE', de, '=UND(A1>1,5;A1<3;A1<>"a;b")', 'UND'],
    ['fr-FR', 'windows-fr_FR', fr, '=ET(A1>1,5;A1<3;A1<>"a;b")', 'ET'],
  ] as const)(
    'canonicalizes custom conditional formulas under %s',
    async (locale, profile, strings, raw, alias) => {
      const workbook = await WorkbookHandle.createDefault();
      const store = createSpreadsheetStore();
      setSelection(store);
      const host = document.createElement('div');
      document.body.appendChild(host);
      try {
        expect(workbook.setSpreadsheetProfileId(profile)).toBe(true);
        const handle = attachConditionalDialog({
          host,
          store,
          strings,
          getWb: () => workbook,
          getLocale: () => locale,
        });
        handle.open({ mode: 'new', kind: 'formula' });
        const formula = inputWithLabel(strings.conditionalDialog.kindFormula);
        formula.value = raw;
        formula.dispatchEvent(new Event('input', { bubbles: true }));
        clickButton('.fc-conddlg__addrow .fc-fmtdlg__btn--primary');
        expect(store.getState().conditional.rules[0]).toMatchObject({
          kind: 'formula',
          formula: '=AND(A1>1.5,A1<3,A1<>"a;b")',
        });
        expect(
          syncConditionalRulesToEngine(workbook, store.getState().conditional.rules, 0),
        ).toEqual({
          written: 1,
          skipped: 0,
        });
        expect(workbook.getConditionalFormats(0)[0]?.formula1).toBe('AND(A1>1.5,A1<3,A1<>"a;b")');
        const loaded = await WorkbookHandle.loadBytes(workbook.save());
        try {
          expect(loaded.getConditionalFormats(0)[0]?.formula1).toBe('AND(A1>1.5,A1<3,A1<>"a;b")');
        } finally {
          loaded.dispose();
        }
        expect(raw).toContain(alias);
        handle.detach();
      } finally {
        workbook.dispose();
      }
    },
  );

  it('canonicalizes German custom validation only when committed and localizes it on reopen', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    setSelection(store);
    const host = document.createElement('div');
    document.body.appendChild(host);
    let locale = 'de-DE';
    expect(workbook.setSpreadsheetProfileId('windows-de_DE')).toBe(true);
    const handle = attachFormatDialog({
      host,
      store,
      strings: de,
      getWb: () => workbook,
      getLocale: () => locale,
    });
    try {
      handle.open('more', { mode: 'dataValidation', focus: 'validation' });
      const kind = [...document.querySelectorAll<HTMLSelectElement>('select')].find((select) =>
        [...select.options].some((option) => option.value === 'custom'),
      );
      if (!kind) throw new Error('custom validation kind select not found');
      kind.value = 'custom';
      kind.dispatchEvent(new Event('change', { bubbles: true }));
      const formula = document.querySelector<HTMLInputElement>(
        `input[aria-label="${de.formatDialog.validationFormula}"]`,
      );
      if (!formula) throw new Error('custom validation formula input not found');
      formula.value = '=UND(A1>1,5;A1<3;A1<>"a;b")';
      formula.dispatchEvent(new Event('input', { bubbles: true }));

      clickButton('.fc-fmtdlg__footer button.fc-fmtdlg__btn--primary');
      const stored = store
        .getState()
        .format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.validation;
      expect(stored).toMatchObject({
        kind: 'custom',
        formula: '=AND(A1>1.5,A1<3,A1<>"a;b")',
      });
      expect(workbook.getValidationsForSheet(0)[0]?.formula1).toBe('AND(A1>1.5,A1<3,A1<>"a;b")');

      locale = 'fr-FR';
      expect(workbook.setSpreadsheetProfileId('windows-fr_FR')).toBe(true);
      handle.open('more', { mode: 'dataValidation', focus: 'validation' });
      const localized = document.querySelector<HTMLInputElement>(
        `input[aria-label="${de.formatDialog.validationFormula}"]`,
      );
      expect(localized?.value).toContain('ET(');
      expect(localized?.value).toContain('1,5');
      expect(localized?.value).toContain('"a;b"');
      handle.close();

      const loaded = await WorkbookHandle.loadBytes(workbook.save());
      try {
        expect(loaded.getValidationsForSheet(0)[0]?.formula1).toBe('AND(A1>1.5,A1<3,A1<>"a;b")');
        const restored = createSpreadsheetStore();
        hydrateValidationsFromEngine(loaded, restored, 0);
        expect(
          restored.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.validation,
        ).toMatchObject({ kind: 'custom', formula: 'AND(A1>1.5,A1<3,A1<>"a;b")' });
        expect(loaded.setSpreadsheetProfileId('windows-fr_FR')).toBe(true);
        const restoredHost = document.createElement('div');
        document.body.appendChild(restoredHost);
        const restoredDialog = attachFormatDialog({
          host: restoredHost,
          store: restored,
          strings: fr,
          getWb: () => loaded,
          getLocale: () => 'fr-FR',
        });
        restoredDialog.open('more', { mode: 'dataValidation', focus: 'validation' });
        const restoredFormula = document.querySelector<HTMLInputElement>(
          `input[aria-label="${fr.formatDialog.validationFormula}"]`,
        );
        expect(restoredFormula?.value).toContain('ET(');
        expect(restoredFormula?.value).toContain('1,5');
        restoredDialog.detach();
      } finally {
        loaded.dispose();
      }
    } finally {
      handle.detach();
      workbook.dispose();
    }
  });
});
