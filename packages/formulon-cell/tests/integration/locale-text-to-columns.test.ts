import { describe, expect, it } from 'vitest';
import { textToColumns } from '../../src/commands/text-to-columns.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../src/store/store.js';

describe('locale-sensitive Text to Columns', () => {
  it.each(['mac', 'windows'] as const)(
    'uses the live %s profile after every language switch',
    async (platform) => {
      const workbook = await WorkbookHandle.createDefault();
      const store = createSpreadsheetStore();
      const addr = { sheet: 0, row: 0, col: 0 };
      const range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
      try {
        expect(workbook.isStub).toBe(false);
        for (const [suffix, number, bool, formula] of [
          ['en_US', '1.5', 'TRUE', '=SUM(1,2)'],
          ['ja_JP', '1.5', 'TRUE', '=SUM(1,2)'],
          ['de_DE', '1,5', 'WAHR', '=SUMME(1;2)'],
          ['fr_FR', '1,5', 'VRAI', '=SOMME(1;2)'],
          ['zh_CN', '1.5', 'TRUE', '=SUM(1,2)'],
          ['ko_KR', '1.5', 'TRUE', '=SUM(1,2)'],
          ['th_TH', '1.5', 'TRUE', '=SUM(1,2)'],
          ['en_US', '1.5', 'TRUE', '=SUM(1,2)'],
        ] as const) {
          expect(workbook.setSpreadsheetProfileId(`${platform}-${suffix}`)).toBe(true);
          workbook.setText(addr, `${number}|${bool}|${formula}`);
          mutators.replaceCells(store, workbook.cells(0));
          expect(textToColumns(store.getState(), store, workbook, range, '|')).toBe(3);
          expect(workbook.getValue(addr)).toEqual({ kind: 'number', value: 1.5 });
          expect(workbook.getValue({ ...addr, col: 1 })).toEqual({ kind: 'bool', value: true });
          expect(workbook.cellFormula({ ...addr, col: 2 })).toBe('=SUM(1,2)');
          expect(workbook.getValue({ ...addr, col: 2 })).toEqual({ kind: 'number', value: 3 });
        }
      } finally {
        workbook.dispose();
      }
    },
  );
});
