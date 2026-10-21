import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { applyCellStyle, cellStyleFallbackFormat } from '../../../src/commands/cell-styles.js';
import { accountingFormatCode, numFmtToFormatCode } from '../../../src/engine/format-writeback.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { spreadsheetProfileForLocale } from '../../../src/mount/locale-profile.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

// Accounting codes and number-style codes per locale, as the desktop
// spreadsheet reports them (decimal and group separators in en-US form).
const EXCEL: Record<string, { comma: string; currency0: string; symbol: string }> = {
  'en-US': {
    comma: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    currency0: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
    symbol: '$',
  },
  'de-DE': {
    comma: '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-',
    currency0: '_-* #,##0\\ "€"_-;\\-* #,##0\\ "€"_-;_-* "-"\\ "€"_-;_-@_-',
    symbol: '€',
  },
  'fr-FR': {
    comma: '_ * #,##0.00_)_ ;_ * \\(#,##0.00\\)_ ;_ * "-"??_)_ ;_ @_ ',
    currency0: '_ * #,##0_)\\ "€"_ ;_ * \\(#,##0\\)\\ "€"_ ;_ * "-"_)\\ "€"_ ;_ @_ ',
    symbol: '€',
  },
  'ja-JP': {
    comma: '#,##0.00;[Red]-#,##0.00',
    currency0: '"¥"#,##0;[Red]"¥"\\-#,##0',
    symbol: '¥',
  },
  'zh-CN': {
    comma: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    currency0: '_("¥"* #,##0_);_("¥"* \\(#,##0\\);_("¥"* "-"_);_(@_)',
    symbol: '¥',
  },
  'ko-KR': {
    comma: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    currency0: '_("₩"* #,##0_);_("₩"* \\(#,##0\\);_("₩"* "-"_);_(@_)',
    symbol: '₩',
  },
  'th-TH': {
    comma: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    currency0: '_("฿"* #,##0_);_("฿"* \\(#,##0\\);_("฿"* "-"_);_(@_)',
    symbol: '฿',
  },
};

describe('accounting formats follow the workbook locale', () => {
  for (const [locale, excel] of Object.entries(EXCEL)) {
    describe(locale, () => {
      let wb: WorkbookHandle;
      beforeAll(async () => {
        wb = await WorkbookHandle.createDefault();
        const profile = spreadsheetProfileForLocale(locale, 'mac');
        if (profile) wb.setSpreadsheetProfileId(profile);
      });
      afterAll(() => wb.dispose());

      it('builds every built-in accounting code from the locale template', () => {
        expect(accountingFormatCode(0, '', wb)).toBe(wb.getNumFmtCode(41));
        expect(accountingFormatCode(0, excel.symbol, wb)).toBe(wb.getNumFmtCode(42));
        expect(accountingFormatCode(2, '', wb)).toBe(wb.getNumFmtCode(43));
        expect(accountingFormatCode(2, excel.symbol, wb)).toBe(wb.getNumFmtCode(44));
      });

      it('writes the Comma and Currency [0] styles as the desktop spreadsheet does', () => {
        const code = (id: 'comma' | 'currency0') =>
          numFmtToFormatCode(cellStyleFallbackFormat(id, 'default', locale).numFmt, wb);
        expect(code('comma')).toBe(excel.comma);
        expect(code('currency0')).toBe(excel.currency0);
      });

      it('applies the Comma style in the workbook locale', () => {
        const store = createSpreadsheetStore();
        const range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
        applyCellStyle(store, null, range, 'comma', { getWorkbook: () => wb });
        const numFmt = store.getState().format.formats.get('0:0:0')?.numFmt;
        expect(numFmtToFormatCode(numFmt, wb)).toBe(excel.comma);
      });
    });
  }

  it('substitutes another currency and decimal places into the template', () => {
    expect(accountingFormatCode(3, '£')).toBe(
      '_("£"* #,##0.000_);_("£"* \\(#,##0.000\\);_("£"* "-"???_);_(@_)',
    );
  });
});
