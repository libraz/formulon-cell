import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { textToColumns } from '../../src/commands/text-to-columns.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

/**
 * Integration: split delimited text across columns. Verifies the WASM-side
 * writes round-trip back through the engine and numeric tokens coerce
 * automatically.
 */
describe('integration: text to columns', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
  });

  afterEach(() => sheet.dispose());

  it('splits comma-delimited text and coerces numeric tokens', () => {
    const { instance, workbook } = sheet;
    workbook.setText({ sheet: 0, row: 0, col: 0 }, 'alpha,1,beta');
    workbook.setText({ sheet: 0, row: 1, col: 0 }, 'gamma,2,delta');
    workbook.recalc();

    const max = textToColumns(
      instance.store.getState(),
      workbook,
      { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 },
      ',',
    );

    expect(max).toBe(3);
    expect(workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'alpha',
    });
    expect(workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 1 });
    expect(workbook.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({
      kind: 'text',
      value: 'beta',
    });
    expect(workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'text',
      value: 'gamma',
    });
    expect(workbook.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'number', value: 2 });
    expect(workbook.getValue({ sheet: 0, row: 1, col: 2 })).toEqual({
      kind: 'text',
      value: 'delta',
    });
  });

  it('leaves cells with fewer than two tokens untouched', () => {
    const { instance, workbook } = sheet;
    workbook.setText({ sheet: 0, row: 0, col: 0 }, 'no-delim');
    workbook.recalc();

    const max = textToColumns(
      instance.store.getState(),
      workbook,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      ',',
    );

    expect(max).toBe(0);
    expect(workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'no-delim',
    });
    expect(workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'blank' });
  });

  it('empty delimiter is a no-op (avoids infinite split)', () => {
    const { instance, workbook } = sheet;
    workbook.setText({ sheet: 0, row: 0, col: 0 }, 'abc');
    workbook.recalc();

    const max = textToColumns(
      instance.store.getState(),
      workbook,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      '',
    );

    expect(max).toBe(0);
    expect(workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'abc' });
  });

  it('skips non-text cells', () => {
    const { instance, workbook } = sheet;
    workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 12);
    workbook.setText({ sheet: 0, row: 1, col: 0 }, 'a,b');
    workbook.recalc();

    const max = textToColumns(
      instance.store.getState(),
      workbook,
      { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 },
      ',',
    );

    expect(max).toBe(2);
    expect(workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 12 });
    expect(workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'text', value: 'a' });
    expect(workbook.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'text', value: 'b' });
  });
});
