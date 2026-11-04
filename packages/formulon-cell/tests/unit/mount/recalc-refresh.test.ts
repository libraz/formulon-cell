import { afterEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

describe('cells a recalc changes reach the grid after a workbook API write', () => {
  let sheet: MountedStubSheet | null = null;
  afterEach(() => sheet?.dispose());

  it('turns a spill anchor into #SPILL! when a blocker is written through setText', async () => {
    const workbook = await WorkbookHandle.createDefault();
    if (!workbook.capabilities.spillInfo) return;
    sheet = await mountStubSheet({ workbook });
    workbook.setFormula({ sheet: 0, row: 1, col: 3 }, '=SEQUENCE(5)');
    await Promise.resolve();
    const cell = () => sheet?.instance.store.getState().data.cells.get('0:1:3')?.value;
    expect(cell()).toEqual({ kind: 'number', value: 1 });
    workbook.setText({ sheet: 0, row: 3, col: 3 }, 'x');
    await Promise.resolve();
    expect(cell()).toMatchObject({ kind: 'error', code: 8 });
    // The blocked spill's other cells are gone from the grid too.
    expect(sheet.instance.store.getState().data.cells.get('0:2:3')?.value.kind ?? 'blank').toBe(
      'blank',
    );
  });
});
