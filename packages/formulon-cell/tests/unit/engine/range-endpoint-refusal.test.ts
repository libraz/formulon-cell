import { FormulonError, StatusCode, type Workbook } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const rawOf = (wb: WorkbookHandle): Workbook => (wb as unknown as { wb: Workbook }).wb;

async function withSheet(
  run: (sheet: ReturnType<Workbook['sheet']>, raw: Workbook) => void | Promise<void>,
): Promise<void> {
  const wb = await WorkbookHandle.createDefault();
  try {
    const sheet = rawOf(wb).sheet(0);
    sheet.setNumber('A2', 1);
    sheet.setNumber('A3', 2);
    sheet.setNumber('A4', 3);
    await run(sheet, rawOf(wb));
  } finally {
    wb.dispose();
  }
}

const refusal = (fn: () => void): FormulonError => {
  try {
    fn();
  } catch (e) {
    expect(e).toBeInstanceOf(FormulonError);
    return e as FormulonError;
  }
  throw new Error('expected setFormula to throw');
};

describe('range operator endpoints', () => {
  it('refuses a function call as the right endpoint', async () => {
    await withSheet((sheet) => {
      const text = '=A9:IFERROR(A2:A4,0)';
      const error = refusal(() => sheet.setFormula('B1', text));
      expect(error.code).toBe(StatusCode.ParserInvalidReference);
      const span = error.parseError;
      expect(span).toBeDefined();
      expect(span?.start).toBeLessThan(span?.end ?? 0);
      expect(span?.end).toBeLessThanOrEqual(text.length);
      expect(text.slice(span?.start, span?.end)).toContain('IFERROR');
    });
  });

  it('refuses a function call as the left endpoint', async () => {
    await withSheet((sheet) => {
      const text = '=SORT(A2:A4):A9';
      const error = refusal(() => sheet.setFormula('B1', text));
      expect(error.code).toBe(StatusCode.ParserInvalidReference);
      const span = error.parseError;
      expect(span).toBeDefined();
      expect(span?.start).toBeLessThan(span?.end ?? 0);
      expect(span?.end).toBeLessThanOrEqual(text.length);
      expect(text.slice(span?.start, span?.end)).toContain('SORT');
    });
  });

  const spilling: [string, string][] = [
    ['LET as the right endpoint', '=A9:LET(a,A2:A4,a)'],
    ['LET as the left endpoint', '=LET(a,A2:A4,a):A9'],
    ['an invoked LAMBDA as the left endpoint', '=LAMBDA(x,x)(A2:A4):A9'],
    ['an implicit-intersection reference as the right endpoint', '=A9:@A2:A4'],
  ];

  for (const [label, text] of spilling) {
    // Waits for the engine to take this endpoint and spill over the bounding area A2:A9.
    it.fails(`spills over the bounding area with ${label}`, async () => {
      await withSheet((sheet, raw) => {
        sheet.setFormula('C1', text);
        raw.recalc();
        const spill = sheet.getSpill('C1');
        expect(spill).not.toBeNull();
        expect(spill?.blocked).toBe(false);
        expect(spill?.range).toMatchObject({ firstRow: 1, lastRow: 8 });
      });
    });
  }
});
