import { afterEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { de, en, fr, ja, ko, th, zh } from '../../src/i18n/strings.js';
import { attachFxDialog } from '../../src/interact/fx-dialog.js';
import { attachWatchPanel } from '../../src/interact/watch-panel.js';
import { createSpreadsheetStore, mutators } from '../../src/store/store.js';

const LOCALES = [
  ['en-US', 'windows-en_US', en, 'English description'],
  ['ja-JP', 'windows-ja_JP', ja, '日本語の説明'],
  ['de-DE', 'windows-de_DE', de, 'Deutsche Beschreibung'],
  ['fr-FR', 'windows-fr_FR', fr, 'Description française'],
  ['zh-CN', 'windows-zh_CN', zh, '中文说明'],
  ['ko-KR', 'windows-ko_KR', ko, '한국어 설명'],
  ['th-TH', 'windows-th_TH', th, 'คำอธิบายภาษาไทย'],
] as const;

afterEach(() => {
  document.body.replaceChildren();
});

describe('seven-language function metadata and Watch Window', () => {
  it.each(LOCALES)('selects provider metadata for %s', async (locale, profile, strings, text) => {
    const workbook = await WorkbookHandle.createDefault();
    const host = document.createElement('div');
    document.body.appendChild(host);
    let currentLocale = locale;
    try {
      expect(workbook.setSpreadsheetProfileId(profile)).toBe(true);
      workbook.setFunctionMetadataProvider({
        ACOS: {
          localized: { [locale]: { description: text } },
        },
      });
      expect(workbook.functionMetadataForLocale('ACOS', locale)?.description).toBe(text);

      const handle = attachFxDialog({
        host,
        strings,
        store: createSpreadsheetStore(),
        getWb: () => workbook,
        getLocale: () => currentLocale,
        onInsert: () => {},
      });
      handle.open();
      const item = document.querySelector<HTMLElement>(
        '[data-fx-name="ACOS"] .fc-fxdialog__item-desc',
      );
      expect(item?.textContent).toBe(text);
      handle.close();

      // The same mounted picker must request the new BCP-47 provider entry
      // after a runtime switch; no stale ordinal-only catalog is acceptable.
      const next = LOCALES.find(([candidate]) => candidate !== locale);
      if (!next) return;
      const [nextLocale, nextProfile, nextStrings, nextText] = next;
      currentLocale = nextLocale;
      expect(workbook.setSpreadsheetProfileId(nextProfile)).toBe(true);
      workbook.setFunctionMetadataProvider({
        ACOS: {
          localized: {
            [locale]: { description: text },
            [nextLocale]: { description: nextText },
          },
        },
      });
      // FxDialog refreshes the locale through the live getter. Its dictionary
      // is presentation-only here; the metadata selection is the assertion.
      void nextStrings;
      handle.refresh();
      handle.open();
      expect(
        document.querySelector<HTMLElement>('[data-fx-name="ACOS"] .fc-fxdialog__item-desc')
          ?.textContent,
      ).toBe(nextText);
      handle.detach();
    } finally {
      workbook.dispose();
    }
  });

  it('keeps watched rows while switching profile display and formula text', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const host = document.createElement('div');
    document.body.appendChild(host);
    let locale = 'en-US';
    const addr = { sheet: 0, row: 0, col: 0 };
    try {
      expect(workbook.setSpreadsheetProfileId('windows-en_US')).toBe(true);
      workbook.setFormula(addr, '=SUM(1.5,2)');
      workbook.recalc();
      mutators.setWatchPanelOpen(store, true);
      mutators.addWatch(store, addr);
      const handle = attachWatchPanel({
        host,
        store,
        getWb: () => workbook,
        getLocale: () => locale,
        workbookName: 'Locale.xlsx',
      });
      const row = (): HTMLTableRowElement => {
        const value = host.querySelector<HTMLTableRowElement>('.fc-watch__row');
        if (!value) throw new Error('watch row missing');
        return value;
      };
      const englishRow = row();
      expect(englishRow.querySelector('td')?.textContent).toBe('Locale.xlsx');
      expect(englishRow.querySelector('.fc-watch__formula')?.textContent).toBe(
        workbook.getFormulaForDisplay(workbook.cellFormula(addr) ?? ''),
      );
      expect(englishRow.querySelector('.fc-watch__value')?.textContent).toBe(
        workbook.getDisplayText(addr.sheet, addr.row, addr.col),
      );

      locale = 'de-DE';
      expect(workbook.setSpreadsheetProfileId('windows-de_DE')).toBe(true);
      workbook.recalc();
      handle.refresh();
      const germanRow = row();
      expect(germanRow.querySelector('.fc-watch__formula')?.textContent).toBe(
        workbook.getFormulaForDisplay(workbook.cellFormula(addr) ?? ''),
      );
      expect(germanRow.querySelector('.fc-watch__value')?.textContent).toBe(
        workbook.getDisplayText(addr.sheet, addr.row, addr.col),
      );
      expect(store.getState().watch.watches).toEqual([addr]);
      handle.detach();
    } finally {
      workbook.dispose();
    }
  });
});
