import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { autoSum } from '../../src/commands/auto-sum.js';
import { computeF9Preview } from '../../src/commands/f9-preview.js';
import {
  applyFlashFill,
  type FlashFillPattern,
  inferFlashFillPattern,
} from '../../src/commands/flash-fill.js';
import {
  clearSparkline,
  clearSparklinesInRange,
  listSparklines,
  setSparkline,
  sparklineAt,
} from '../../src/commands/sparkline.js';
import { mutators } from '../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

/**
 * Integration: a handful of commands that have solid unit coverage but no
 * end-to-end check against the live mount. The unit suites pin the
 * algorithms; this file verifies that the store/workbook round-trips behave
 * as the chrome layer expects when these commands are dispatched from the
 * mounted sheet.
 */
describe('integration: autoSum against the live mount', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
  });

  afterEach(() => sheet.dispose());

  it('writes a column SUM under a contiguous numeric block and recalcs to the right value', () => {
    const { instance, workbook } = sheet;
    for (let r = 0; r < 4; r += 1) {
      workbook.setNumber({ sheet: 0, row: r, col: 0 }, r + 1);
    }
    workbook.recalc();
    mutators.replaceCells(instance.store, workbook.cells(0));
    mutators.setActive(instance.store, { sheet: 0, row: 4, col: 0 });

    const result = autoSum(instance.store.getState(), workbook);
    expect(result).toEqual({ addr: { sheet: 0, row: 4, col: 0 }, formula: '=SUM(A1:A4)' });

    workbook.recalc();
    mutators.replaceCells(instance.store, workbook.cells(0));
    const total = workbook.getValue({ sheet: 0, row: 4, col: 0 });
    expect(total.kind === 'number' ? total.value : null).toBe(10);
  });

  it('places a range SUM directly below a selected block', () => {
    const { instance, workbook } = sheet;
    for (let r = 0; r < 3; r += 1) {
      for (let c = 0; c < 2; c += 1) {
        workbook.setNumber({ sheet: 0, row: r, col: c }, r + c);
      }
    }
    workbook.recalc();
    mutators.replaceCells(instance.store, workbook.cells(0));
    mutators.setRange(instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 });

    // No empty trailing rows inside the selection ⇒ falls through to a
    // single `=SUM(<range>)` placed directly below the block.
    const result = autoSum(instance.store.getState(), workbook);
    expect(result?.addr).toEqual({ sheet: 0, row: 3, col: 0 });
    expect(result?.formula).toBe('=SUM(A1:B3)');
  });

  it('returns null when there is no adjacent numeric block', () => {
    const { instance, workbook } = sheet;
    mutators.setActive(instance.store, { sheet: 0, row: 4, col: 4 });
    expect(autoSum(instance.store.getState(), workbook)).toBeNull();
  });
});

describe('integration: F9 preview against live cells', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
    const { instance, workbook } = sheet;
    workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 42);
    workbook.setText({ sheet: 0, row: 0, col: 1 }, 'hello');
    workbook.recalc();
    mutators.replaceCells(instance.store, workbook.cells(0));
  });

  afterEach(() => sheet.dispose());

  it('resolves an A1 ref against the live cell map', () => {
    const cells = sheet.instance.store.getState().data.cells;
    const out = computeF9Preview('=A1+B1', 'A1', 0, cells);
    expect(out).toEqual({ display: '42', substitutable: true });
  });

  it('returns the text rendering of a text-valued ref', () => {
    const cells = sheet.instance.store.getState().data.cells;
    const out = computeF9Preview('=B1', 'B1', 0, cells);
    expect(out).toEqual({ display: '"hello"', substitutable: true });
  });

  it('reports unsupported for a sub-expression', () => {
    const cells = sheet.instance.store.getState().data.cells;
    expect(computeF9Preview('=A1+B1', 'A1+B1', 0, cells)).toEqual({
      display: '',
      substitutable: false,
    });
  });

  it('reports #REF! when sheet-prefixed and the sheet is unknown', () => {
    const cells = sheet.instance.store.getState().data.cells;
    const out = computeF9Preview('=Other!A1', 'Other!A1', 0, cells, () => -1);
    expect(out).toEqual({ display: '#REF!', substitutable: false });
  });
});

describe('integration: flash fill applied across pending inputs', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
  });

  afterEach(() => sheet.dispose());

  it('inferred token-pattern fills a target column from a source column', () => {
    const { instance, workbook } = sheet;
    const examples = [{ input: 'John Smith', output: 'John' }];
    const pattern = inferFlashFillPattern(examples);
    expect(pattern).toEqual({ kind: 'token', delimiter: ' ', index: 0 });

    const pendingInputs = ['Jane Doe', 'Adam Stone'];
    const filled = applyFlashFill(pattern as FlashFillPattern, pendingInputs);
    expect(filled).toEqual(['Jane', 'Adam']);

    // The mount surface for flash-fill writes via wb.setText — exercise that
    // path so we know the round-trip lands in the store.
    pendingInputs.forEach((_input, i) => {
      const out = filled[i];
      if (out !== null && out !== undefined) {
        workbook.setText({ sheet: 0, row: i + 1, col: 1 }, out);
      }
    });
    mutators.replaceCells(instance.store, workbook.cells(0));
    expect(workbook.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({
      kind: 'text',
      value: 'Jane',
    });
    expect(workbook.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({
      kind: 'text',
      value: 'Adam',
    });
  });

  it('returns null for inputs that fail the pattern (slice out-of-range)', () => {
    const pattern: FlashFillPattern = { kind: 'substring', start: 5, length: 3 };
    const out = applyFlashFill(pattern, ['hi', 'longstring']);
    expect(out[0]).toBeNull();
    // longstring[5..8] → "tri"
    expect(out[1]).toBe('tri');
  });

  it('infers a casing transform from a single example', () => {
    const pattern = inferFlashFillPattern([{ input: 'hello', output: 'HELLO' }]);
    expect(pattern).toEqual({ kind: 'case', mode: 'upper' });
  });
});

describe('integration: sparkline lifecycle against the mounted store', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
  });

  afterEach(() => sheet.dispose());

  it('setSparkline persists in store.sparkline and is readable by addr', () => {
    const { instance } = sheet;
    setSparkline(instance.store, { sheet: 0, row: 5, col: 2 }, { kind: 'line', source: 'A1:E1' });
    const at = sparklineAt(instance.store.getState(), { sheet: 0, row: 5, col: 2 });
    expect(at).toEqual({ kind: 'line', source: 'A1:E1' });
  });

  it('listSparklines returns entries sorted by sheet/row/col', () => {
    const { instance } = sheet;
    setSparkline(instance.store, { sheet: 0, row: 1, col: 1 }, { kind: 'column', source: 'A1:C1' });
    setSparkline(instance.store, { sheet: 0, row: 0, col: 5 }, { kind: 'line', source: 'A1:C1' });
    const all = listSparklines(instance.store.getState());
    expect(all.map((e) => `${e.addr.row}:${e.addr.col}`)).toEqual(['0:5', '1:1']);
  });

  it('clearSparkline removes a single entry; clearSparklinesInRange clears overlap', () => {
    const { instance } = sheet;
    for (let c = 0; c < 3; c += 1) {
      setSparkline(instance.store, { sheet: 0, row: 0, col: c }, { kind: 'line', source: 'A2:E2' });
    }
    clearSparkline(instance.store, { sheet: 0, row: 0, col: 1 });
    expect(listSparklines(instance.store.getState()).map((e) => e.addr.col)).toEqual([0, 2]);

    clearSparklinesInRange(instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 4 });
    expect(listSparklines(instance.store.getState())).toEqual([]);
  });
});
