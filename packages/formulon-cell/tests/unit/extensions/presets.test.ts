import { describe, expect, it } from 'vitest';
import { ALL_FEATURE_IDS, resolveFlags } from '../../../src/extensions/features.js';
import { full, minimal, standard } from '../../../src/extensions/presets.js';

describe('feature presets', () => {
  it('minimal keeps only the bare spreadsheet chrome', () => {
    expect(minimal()).toMatchObject({
      viewToolbar: false,
      workbookObjects: false,
      quickAnalysis: false,
      charts: false,
      pivotTableDialog: false,
      sheetTabs: false,
      contextMenu: false,
      iterative: false,
      gotoSpecial: false,
      pageSetup: false,
      errorIndicators: false,
    });
  });

  it('standard keeps lightweight chrome but suppresses authoring dialogs', () => {
    const p = standard();
    expect(p.viewToolbar).toBeUndefined();
    expect(p.workbookObjects).toBeUndefined();
    expect(p.quickAnalysis).toBeUndefined();
    expect(p.charts).toBeUndefined();
    expect(p.pivotTableDialog).toBe(false);
    expect(p.formatDialog).toBe(false);
    expect(p.conditional).toBe(false);
    expect(p.iterative).toBe(false);
    expect(p.gotoSpecial).toBe(false);
    expect(p.pageSetup).toBe(false);
    expect(p.namedRanges).toBe(false);
  });

  it('full remains the default full feature surface', () => {
    expect(full()).toEqual({});
  });

  it('resolves feature defaults with only heavy panels default-off', () => {
    const flags = resolveFlags();
    expect(flags.watchWindow).toBe(false);
    expect(flags.slicer).toBe(false);
    for (const id of ALL_FEATURE_IDS) {
      if (id === 'watchWindow' || id === 'slicer') continue;
      expect(flags[id]).toBe(true);
    }
  });

  it('allows presets to explicitly opt out of default-on features', () => {
    const flags = resolveFlags(minimal());
    expect(flags.formulaBar).toBe(true);
    expect(flags.statusBar).toBe(true);
    expect(flags.shortcuts).toBe(true);
    expect(flags.viewToolbar).toBe(false);
    expect(flags.quickAnalysis).toBe(false);
    expect(flags.charts).toBe(false);
    expect(flags.iterative).toBe(false);
  });
});
