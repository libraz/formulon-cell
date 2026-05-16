import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { setCellLocked, setProtectedSheet } from '../../src/commands/protection.js';
import { textToColumns } from '../../src/commands/text-to-columns.js';
import { mutators } from '../../src/store/store.js';
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
      instance.store,
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

  it('can treat consecutive delimiters as one', () => {
    const { instance, workbook } = sheet;
    workbook.setText({ sheet: 0, row: 0, col: 0 }, 'alpha,,1,,,beta');
    workbook.recalc();

    const max = textToColumns(
      instance.store.getState(),
      instance.store,
      workbook,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      ',',
      { collapseConsecutiveDelimiters: true },
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
  });

  it('leaves cells with fewer than two tokens untouched', () => {
    const { instance, workbook } = sheet;
    workbook.setText({ sheet: 0, row: 0, col: 0 }, 'no-delim');
    workbook.recalc();

    const max = textToColumns(
      instance.store.getState(),
      instance.store,
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
      instance.store,
      workbook,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      '',
    );

    expect(max).toBe(0);
    expect(workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'abc' });
  });

  it('splits by multiple delimiters for the ribbon dialog path', () => {
    const { instance, workbook } = sheet;
    workbook.setText({ sheet: 0, row: 0, col: 0 }, 'alpha,1;beta gamma');
    workbook.recalc();

    const max = textToColumns(
      instance.store.getState(),
      instance.store,
      workbook,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      [',', ';', ' '],
    );

    expect(max).toBe(4);
    expect(workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'alpha',
    });
    expect(workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'number', value: 1 });
    expect(workbook.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({
      kind: 'text',
      value: 'beta',
    });
    expect(workbook.getValue({ sheet: 0, row: 0, col: 3 })).toEqual({
      kind: 'text',
      value: 'gamma',
    });
  });

  it('skips non-text cells', () => {
    const { instance, workbook } = sheet;
    workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 12);
    workbook.setText({ sheet: 0, row: 1, col: 0 }, 'a,b');
    workbook.recalc();

    const max = textToColumns(
      instance.store.getState(),
      instance.store,
      workbook,
      { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 },
      ',',
    );

    expect(max).toBe(2);
    expect(workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 12 });
    expect(workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'text', value: 'a' });
    expect(workbook.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'text', value: 'b' });
  });

  it('preserves split tokens as text when destination cells are Text formatted', () => {
    const { instance, workbook } = sheet;
    workbook.setText({ sheet: 0, row: 0, col: 0 }, '00123,=A1');
    workbook.recalc();
    mutators.setCellFormat(
      instance.store,
      { sheet: 0, row: 0, col: 0 },
      { numFmt: { kind: 'text' } },
    );
    mutators.setCellFormat(
      instance.store,
      { sheet: 0, row: 0, col: 1 },
      { numFmt: { kind: 'text' } },
    );

    const max = textToColumns(
      instance.store.getState(),
      instance.store,
      workbook,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      ',',
    );

    expect(max).toBe(2);
    expect(workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: '00123',
    });
    expect(workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'text',
      value: '=A1',
    });
  });

  it('copies the source cell format to split destination cells', () => {
    const { instance, workbook } = sheet;
    workbook.setText({ sheet: 0, row: 0, col: 0 }, 'alpha,beta');
    workbook.recalc();
    mutators.setCellFormat(instance.store, { sheet: 0, row: 0, col: 0 }, { fill: '#c6efce' });

    const max = textToColumns(
      instance.store.getState(),
      instance.store,
      workbook,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      ',',
    );

    expect(max).toBe(2);
    expect(instance.store.getState().format.formats.get('0:0:0')).toEqual({ fill: '#c6efce' });
    expect(instance.store.getState().format.formats.get('0:0:1')).toEqual({ fill: '#c6efce' });
  });

  it('skips locked protected destinations while writing unlocked split targets', () => {
    const { instance, workbook } = sheet;
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    workbook.setText({ sheet: 0, row: 0, col: 0 }, 'alpha,1,beta');
    workbook.setText({ sheet: 0, row: 0, col: 1 }, 'locked');
    workbook.recalc();
    setCellLocked(instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, false);
    setCellLocked(instance.store, { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 }, false);
    setProtectedSheet(instance.store, 0, true);

    try {
      const max = textToColumns(
        instance.store.getState(),
        instance.store,
        workbook,
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        ',',
      );

      expect(max).toBe(3);
      expect(workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
        kind: 'text',
        value: 'alpha',
      });
      expect(workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
        kind: 'text',
        value: 'locked',
      });
      expect(workbook.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({
        kind: 'text',
        value: 'beta',
      });
      expect(warn).toHaveBeenCalledTimes(1);
    } finally {
      warn.mockRestore();
    }
  });

  it('leaves locked protected source cells unchanged', () => {
    const { instance, workbook } = sheet;
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    workbook.setText({ sheet: 0, row: 0, col: 0 }, 'alpha,1');
    workbook.recalc();
    setCellLocked(instance.store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, false);
    setProtectedSheet(instance.store, 0, true);

    try {
      const max = textToColumns(
        instance.store.getState(),
        instance.store,
        workbook,
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        ',',
      );

      expect(max).toBe(2);
      expect(workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
        kind: 'text',
        value: 'alpha,1',
      });
      expect(workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
        kind: 'number',
        value: 1,
      });
      expect(warn).toHaveBeenCalledTimes(1);
    } finally {
      warn.mockRestore();
    }
  });
});
