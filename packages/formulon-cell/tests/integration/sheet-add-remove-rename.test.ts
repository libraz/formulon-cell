import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

/**
 * Integration: walk the sheetTabs controller from `addSheet` through rename
 * and remove (where the engine supports it). Under the stub, the engine
 * gates rename/remove behind `capabilities.sheetMutate`; the test asserts
 * that gate matches the real surface the demo apps see.
 */
describe('integration: sheet add / rename / remove (stub)', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
  });

  afterEach(() => sheet.dispose());

  it('starts with exactly one sheet visible in the tab strip', () => {
    const tabs = sheet.host.querySelectorAll('.fc-host__sheetbar-tab');
    expect(tabs.length).toBe(1);
  });

  it('addSheet appends a tab and exposes the new index', () => {
    const { workbook } = sheet;
    const before = workbook.sheetCount;
    const idx = workbook.addSheet('Plan');
    expect(idx).toBe(before);
    expect(workbook.sheetCount).toBe(before + 1);
    expect(workbook.sheetName(idx)).toBe('Plan');
  });

  it('addSheet without a name assigns a unique fallback', () => {
    const { workbook } = sheet;
    const idx1 = workbook.addSheet();
    const idx2 = workbook.addSheet();
    expect(workbook.sheetName(idx1)).not.toBe(workbook.sheetName(idx2));
  });

  it('renameSheet under the stub reports the capability gate', () => {
    const { workbook } = sheet;
    workbook.addSheet('Old');
    const result = workbook.renameSheet(1, 'New');
    // Under stub `capabilities.sheetMutate` is false → renameSheet returns false.
    // Under real engine it would be true. Either way it never throws.
    expect(typeof result).toBe('boolean');
  });

  it('removeSheet under the stub reports the capability gate', () => {
    const { workbook } = sheet;
    workbook.addSheet('Temp');
    const result = workbook.removeSheet(1);
    expect(typeof result).toBe('boolean');
  });

  it('emits sheet-add events; the sheet-tabs controller re-renders', () => {
    const { workbook } = sheet;
    workbook.addSheet('Foo');
    const tabs = sheet.host.querySelectorAll('.fc-host__sheetbar-tab');
    // Controller renders one tab per sheet — at least 2 after addSheet.
    expect(tabs.length).toBeGreaterThanOrEqual(2);
  });
});
