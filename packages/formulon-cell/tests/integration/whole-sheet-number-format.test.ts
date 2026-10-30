import { describe, expect, it } from 'vitest';
import { MAX_COL, MAX_ROW } from '../../src/engine/address.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { mutators } from '../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/mount.js';

const cells = [
  { sheet: 0, row: 0, col: 0 },
  { sheet: 0, row: 1, col: 0 },
  { sheet: 0, row: 2, col: 0 },
  { sheet: 0, row: 3, col: 0 },
  { sheet: 0, row: 4, col: 0 },
] as const;
const far = { sheet: 0, row: MAX_ROW, col: MAX_COL };

function numberFormatAt(workbook: WorkbookHandle, addr: (typeof cells)[number] | typeof far) {
  return workbook.getEffectiveStyle(addr.sheet, addr.row, addr.col)?.numFmtCode;
}

async function mountMixedWorkbook(): Promise<MountedStubSheet> {
  const workbook = await WorkbookHandle.createDefault({ locale: 'en' });
  expect(workbook.isStub).toBe(false);
  workbook.setNumber(cells[0], 1234.5);
  workbook.setText(cells[1], '00123');
  workbook.setText(cells[2], 'abc');
  workbook.setFormula(cells[3], '=SUM(A1,1)');
  workbook.setFormula(cells[4], '="label"');
  workbook.recalc();

  const mounted = await mountStubSheet({ workbook, locale: 'en' });
  const formats = [
    { kind: 'currency', symbol: '$', decimals: 0 },
    { kind: 'date', pattern: 'm/d/yyyy' },
    { kind: 'percent', decimals: 1 },
    { kind: 'scientific', decimals: 3 },
    { kind: 'custom', pattern: '0.000' },
  ] as const;
  for (const [index, numFmt] of formats.entries()) {
    const addr = cells[index];
    if (!addr) throw new Error(`Missing seeded cell ${index}`);
    mutators.setCellFormat(mounted.instance.store, addr, { numFmt });
  }
  return mounted;
}

function applyFixedTwoDecimalsToWholeSheet(mounted: MountedStubSheet): void {
  const { instance } = mounted;
  mutators.selectAll(instance.store);
  instance.openFormatDialog('number');
  const dialog = document.querySelector<HTMLElement>('[class="fc-fmtdlg"]');
  if (!dialog) throw new Error('Format Cells dialog did not open');

  const fixed = dialog.querySelector<HTMLButtonElement>('button[data-fc-cat="fixed"]');
  if (!fixed) throw new Error('Format Cells Number category is missing');
  fixed.click();

  const decimals = dialog.querySelector<HTMLInputElement>(
    '.fc-fmtdlg__cat-controls input[type="number"]',
  );
  if (!decimals) throw new Error('Format Cells decimal places input is missing');
  decimals.value = '2';
  decimals.dispatchEvent(new Event('input', { bubbles: true }));

  const ok = dialog.querySelector<HTMLButtonElement>('.fc-fmtdlg__footer .fc-fmtdlg__btn--primary');
  if (!ok) throw new Error('Format Cells OK button is missing');
  ok.click();
}

const valuesAndFormulas = (workbook: WorkbookHandle) =>
  cells.map((addr) => ({
    value: workbook.getValue(addr),
    formula: workbook.cellFormula(addr),
  }));

describe('real WASM whole-sheet number format', () => {
  it('formats numbers and numeric formulas while leaving text values intact, sparse, and persistent', async () => {
    const mounted = await mountMixedWorkbook();
    const { instance, workbook } = mounted;
    try {
      const before = valuesAndFormulas(workbook);
      expect(before).toEqual([
        { value: { kind: 'number', value: 1234.5 }, formula: null },
        { value: { kind: 'text', value: '00123' }, formula: null },
        { value: { kind: 'text', value: 'abc' }, formula: null },
        { value: { kind: 'number', value: 1235.5 }, formula: '=SUM(A1,1)' },
        { value: { kind: 'text', value: 'label' }, formula: '="label"' },
      ]);
      const oldFormats = cells.map((addr) => numberFormatAt(workbook, addr));
      expect(new Set(oldFormats).size).toBe(5);

      applyFixedTwoDecimalsToWholeSheet(mounted);

      expect(valuesAndFormulas(workbook)).toEqual(before);
      expect(cells.map((addr) => numberFormatAt(workbook, addr))).toEqual(
        Array(cells.length).fill('0.00'),
      );
      expect(numberFormatAt(workbook, far)).toBe('0.00');
      expect(workbook.getDisplayText(0, 0, 0)).toBe('1234.50');
      expect(workbook.getDisplayText(0, 3, 0)).toBe('1235.50');
      expect(workbook.getDisplayText(0, 1, 0)).toBe('00123');
      expect(workbook.getDisplayText(0, 2, 0)).toBe('abc');
      expect(workbook.getDisplayText(0, 4, 0)).toBe('label');
      expect(instance.store.getState().data.cells.size).toBe(cells.length);
      expect(instance.store.getState().format.formats.size).toBeLessThan(20);

      const future = { sheet: 0, row: 0, col: 1 };
      expect(instance.applyChanges([{ addr: future, input: '42.5' }]).status).toBe('applied');
      expect(workbook.getValue(future)).toEqual({ kind: 'number', value: 42.5 });
      expect(numberFormatAt(workbook, future)).toBe('0.00');
      expect(workbook.getDisplayText(future.sheet, future.row, future.col)).toBe('42.50');
      expect(instance.store.getState().data.cells.size).toBe(cells.length + 1);
      expect(instance.store.getState().format.formats.size).toBeLessThan(20);

      const reloaded = await WorkbookHandle.loadBytes(workbook.save());
      const reloadedMount = await mountStubSheet({ workbook: reloaded, locale: 'en' });
      try {
        expect(valuesAndFormulas(reloaded)).toEqual(before);
        expect(reloaded.getValue(future)).toEqual({ kind: 'number', value: 42.5 });
        expect(cells.map((addr) => numberFormatAt(reloaded, addr))).toEqual(
          Array(cells.length).fill('0.00'),
        );
        expect(numberFormatAt(reloaded, future)).toBe('0.00');
        expect(numberFormatAt(reloaded, far)).toBe('0.00');
        expect(reloaded.getDisplayText(0, 1, 0)).toBe('00123');
        expect(reloaded.getDisplayText(0, 2, 0)).toBe('abc');
        expect(reloaded.getDisplayText(0, 4, 0)).toBe('label');
        const afterReload = { sheet: 0, row: 1, col: 1 };
        expect(
          reloadedMount.instance.applyChanges([{ addr: afterReload, input: '8.25' }]).status,
        ).toBe('applied');
        expect(reloaded.getValue(afterReload)).toEqual({ kind: 'number', value: 8.25 });
        expect(numberFormatAt(reloaded, afterReload)).toBe('0.00');
      } finally {
        reloadedMount.dispose();
        reloaded.dispose();
      }
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('undo restores the preexisting heterogeneous formats without changing data or formulas', async () => {
    const mounted = await mountMixedWorkbook();
    const { instance, workbook } = mounted;
    try {
      const beforeValues = valuesAndFormulas(workbook);
      const beforeFormats = cells.map((addr) => numberFormatAt(workbook, addr));
      applyFixedTwoDecimalsToWholeSheet(mounted);
      expect(cells.map((addr) => numberFormatAt(workbook, addr))).toEqual(
        Array(cells.length).fill('0.00'),
      );

      expect(instance.undo()).toBe(true);
      expect(valuesAndFormulas(workbook)).toEqual(beforeValues);
      expect(cells.map((addr) => numberFormatAt(workbook, addr))).toEqual(beforeFormats);
      expect(numberFormatAt(workbook, far)).not.toBe('0.00');
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });
});
