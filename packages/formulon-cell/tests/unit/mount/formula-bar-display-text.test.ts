import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

describe('formula bar display text', () => {
  let sheet: MountedStubSheet;
  let fx: HTMLTextAreaElement;
  const at = (row: number, col: number) => ({ sheet: 0, row, col });

  const select = (row: number, col: number): void => {
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, at(row, col));
  };

  beforeEach(async () => {
    sheet = await mountStubSheet();
    fx = sheet.host.querySelector('.fc-host__formulabar-input') as HTMLTextAreaElement;
  });
  afterEach(() => sheet.dispose());

  it('keeps the number format for dates, times and percentages', () => {
    sheet.workbook.setNumber(at(0, 0), 45296);
    sheet.workbook.setNumber(at(1, 0), 0.256);
    sheet.workbook.setNumber(at(2, 0), 13.75 / 24 + 30 / 86400);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    const store = sheet.instance.store;
    mutators.setCellFormat(store, at(0, 0), { numFmt: { kind: 'date', pattern: 'yyyy/m/d' } });
    mutators.setCellFormat(store, at(1, 0), { numFmt: { kind: 'percent', decimals: 1 } });
    mutators.setCellFormat(store, at(2, 0), { numFmt: { kind: 'time', pattern: 'h:mm:ss' } });
    select(0, 0);
    expect(fx.value).toBe('2024/1/5');
    select(1, 0);
    expect(fx.value).toBe('25.6%');
    select(2, 0);
    expect(fx.value).toBe('13:45:30');
  });

  it('treats custom date patterns like dates', () => {
    sheet.workbook.setNumber(at(0, 0), 45296);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setCellFormat(sheet.instance.store, at(0, 0), {
      numFmt: { kind: 'custom', pattern: 'yyyy"年"m"月"d"日"' },
    });
    select(0, 0);
    expect(fx.value).toBe('2024年1月5日');
  });

  it('shows the raw number for currency and scientific formats', () => {
    sheet.workbook.setNumber(at(0, 0), 1234.5);
    sheet.workbook.setNumber(at(1, 0), 12345678901234);
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setCellFormat(sheet.instance.store, at(0, 0), {
      numFmt: { kind: 'currency', decimals: 0, symbol: '¥' },
    });
    mutators.setCellFormat(sheet.instance.store, at(1, 0), {
      numFmt: { kind: 'scientific', decimals: 2 },
    });
    select(0, 0);
    expect(fx.value).toBe('1234.5');
    select(1, 0);
    expect(fx.value).toBe('12345678901234');
  });

  it('keeps embedded line feeds and long text in full', () => {
    const long = 'x'.repeat(300);
    sheet.workbook.setText(at(0, 0), 'a\nb');
    sheet.workbook.setText(at(1, 0), long);
    select(0, 0);
    expect(fx.value).toBe('a\nb');
    select(1, 0);
    expect(fx.value).toBe(long);
  });

  it('shows the anchor formula on dynamic spill members', () => {
    sheet.workbook.setFormula(at(0, 4), '=SEQUENCE(3)');
    sheet.workbook.setNumber(at(1, 4), 2);
    sheet.workbook.setNumber(at(2, 4), 3);
    vi.spyOn(sheet.workbook, 'spillInfo').mockImplementation((_s, row, col) =>
      col === 4 && row <= 2 ? { anchorRow: 0, anchorCol: 4, rows: 3, cols: 1 } : null,
    );
    select(0, 4);
    const anchor = fx.value;
    expect(anchor).toBe('=SEQUENCE(3)');
    select(1, 4);
    expect(fx.value).toBe(anchor);
    select(2, 4);
    expect(fx.value).toBe(anchor);
  });
});
