import { mount, type SpreadsheetInstance } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it } from 'vitest';
import { spreadsheetInternals } from '../../src/mount/internals.js';
import { makeHost } from './helpers.js';

const live: Array<{ dispose(): void }> = [];

const mountSheet = async (): Promise<SpreadsheetInstance> => {
  const instance = await mount(makeHost(), { locale: 'en' });
  live.push(instance);
  return instance;
};

afterEach(() => {
  for (const d of live.splice(0).reverse()) d.dispose();
});

const layoutOf = (instance: SpreadsheetInstance) =>
  spreadsheetInternals(instance).store.getState().layout;
const mergesOf = (instance: SpreadsheetInstance) =>
  spreadsheetInternals(instance).store.getState().merges;

describe('facade journal layout', () => {
  it('round-trips setRowHeight', async () => {
    const inst = await mountSheet();
    const s = inst.workbook.sheet(0);
    expect(s.listRowOverrides().some((r) => r.row === 2)).toBe(false);
    s.setRowHeight(2, 2, 40);
    expect(s.listRowOverrides().find((r) => r.row === 2)?.heightPt).toBe(40);
    expect(layoutOf(inst).rowHeights.has(2)).toBe(true);
    inst.history.undo();
    expect(s.listRowOverrides().find((r) => r.row === 2)?.heightPt ?? null).toBeNull();
    expect(layoutOf(inst).rowHeights.has(2)).toBe(false);
    inst.history.redo();
    expect(s.listRowOverrides().find((r) => r.row === 2)?.heightPt).toBe(40);
    expect(layoutOf(inst).rowHeights.has(2)).toBe(true);
  });

  it('round-trips setColumnHidden', async () => {
    const inst = await mountSheet();
    const s = inst.workbook.sheet(0);
    s.setColumnHidden(1, 1, true);
    expect(layoutOf(inst).hiddenCols.has(1)).toBe(true);
    inst.history.undo();
    expect(s.listColumnOverrides().some((c) => c.hidden)).toBe(false);
    expect(layoutOf(inst).hiddenCols.has(1)).toBe(false);
    inst.history.redo();
    expect(s.listColumnOverrides().some((c) => c.hidden && c.first <= 1 && c.last >= 1)).toBe(true);
    expect(layoutOf(inst).hiddenCols.has(1)).toBe(true);
  });

  it('round-trips addMerge', async () => {
    const inst = await mountSheet();
    const s = inst.workbook.sheet(0);
    s.addMerge('B2:C3');
    expect(s.listMerges()).toHaveLength(1);
    expect(mergesOf(inst).byAnchor.size).toBe(1);
    inst.history.undo();
    expect(s.listMerges()).toHaveLength(0);
    expect(mergesOf(inst).byAnchor.size).toBe(0);
    inst.history.redo();
    expect(s.listMerges()).toHaveLength(1);
    expect(mergesOf(inst).byAnchor.size).toBe(1);
  });

  it('round-trips setRowOutline', async () => {
    const inst = await mountSheet();
    const s = inst.workbook.sheet(0);
    s.setRowOutline(3, 4, 2);
    expect(layoutOf(inst).outlineRows.get(3)).toBe(2);
    inst.history.undo();
    expect(s.listRowOverrides().some((r) => r.outlineLevel > 0)).toBe(false);
    expect(layoutOf(inst).outlineRows.get(3) ?? 0).toBe(0);
    inst.history.redo();
    expect(s.listRowOverrides().filter((r) => r.outlineLevel === 2)).toHaveLength(2);
    expect(layoutOf(inst).outlineRows.get(4)).toBe(2);
  });

  it('restores a prior row state when a later write replaces it', async () => {
    const inst = await mountSheet();
    const s = inst.workbook.sheet(0);
    s.setRowHeight(1, 1, 30);
    s.setRowHidden(1, 1, true);
    inst.history.undo();
    const row = s.listRowOverrides().find((r) => r.row === 1);
    expect(row?.heightPt).toBe(30);
    expect(row?.hidden).toBe(false);
  });
});
