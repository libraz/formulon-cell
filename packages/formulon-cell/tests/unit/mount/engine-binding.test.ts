import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { presets } from '../../../src/extensions/presets.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

/**
 * Unit: engine-binding — the central wiring layer between the workbook
 * handle, store, renderer, and feature handles. Drives the keyboard /
 * clipboard / pointer / context-menu / find-replace / validation /
 * quick-analysis attachments, all gated on resolved feature flags. We
 * exercise it through the mounted sheet so the flag → feature visibility
 * round-trip is checked end-to-end.
 */
describe('mount/engine-binding — feature gating against preset.full()', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
  });

  afterEach(() => sheet.dispose());

  it('attaches clipboard / context-menu / find-replace / validation / quick-analysis', () => {
    const f = sheet.instance.features;
    expect(f.clipboard).toBeTruthy();
    expect(f.contextMenu).toBeTruthy();
    expect(f.findReplace).toBeTruthy();
    expect(f.validation).toBeTruthy();
    expect(f.quickAnalysis).toBeTruthy();
  });

  it('paste-special is attached when both pasteSpecial flag and clipboard are on', () => {
    expect(sheet.instance.features.pasteSpecial).toBeTruthy();
  });
});

describe('mount/engine-binding — feature gating against preset.minimal()', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet({ features: presets.minimal() });
  });

  afterEach(() => sheet.dispose());

  it('does NOT attach context-menu / find-replace / validation / quick-analysis', () => {
    const f = sheet.instance.features;
    expect(f.contextMenu).toBeFalsy();
    expect(f.findReplace).toBeFalsy();
    expect(f.validation).toBeFalsy();
    expect(f.quickAnalysis).toBeFalsy();
  });

  it('paste-special is not attached when its own flag is off', () => {
    expect(sheet.instance.features.pasteSpecial).toBeFalsy();
  });
});

describe('mount/engine-binding — workbook subscribe forwards events', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
  });

  afterEach(() => sheet.dispose());

  it('a wb.setNumber updates store.data.cells via the subscription', () => {
    const a = { sheet: 0, row: 0, col: 0 };
    sheet.workbook.setNumber(a, 99);
    const cells = sheet.instance.store.getState().data.cells;
    const k = `${a.sheet}:${a.row}:${a.col}`;
    expect(cells.get(k)?.value).toEqual({ kind: 'number', value: 99 });
  });

  it('emits `cellChange` on a wb mutation', () => {
    const onChange = vi.fn();
    sheet.instance.on('cellChange', onChange);
    sheet.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 5);
    expect(onChange).toHaveBeenCalledTimes(1);
    const arg = onChange.mock.calls[0]?.[0];
    expect(arg?.value).toEqual({ kind: 'number', value: 5 });
  });

  it('cellChange payload echoes the formula text when set via wb.setFormula', () => {
    const onChange = vi.fn();
    sheet.instance.on('cellChange', onChange);
    sheet.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    sheet.workbook.setFormula({ sheet: 0, row: 0, col: 1 }, '=A1*2');
    sheet.workbook.recalc();
    // Last call is the formula write — wb echoes the source formula in the payload.
    const last = onChange.mock.calls.at(-1)?.[0];
    expect(last?.formula).toBe('=A1*2');
  });
});

describe('mount/engine-binding — grid double-click begins inline edit', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
  });

  afterEach(() => sheet.dispose());

  it('dblclick on the grid opens the inline editor seeded with the active cell', () => {
    sheet.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 7);
    sheet.workbook.recalc();
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });

    const grid = sheet.host.querySelector('.fc-host__grid') as HTMLElement;
    grid.dispatchEvent(new MouseEvent('dblclick', { button: 0, bubbles: true }));

    // The inline-edit textarea is mounted under the host; it should exist
    // with the cell's current rendering as its value.
    const editor = sheet.host.querySelector('.fc-host__editor') as HTMLTextAreaElement | null;
    expect(editor).not.toBeNull();
    expect(editor?.value).toBe('7');
  });

  it('dblclick is ignored while the editor is already active', () => {
    const grid = sheet.host.querySelector('.fc-host__grid') as HTMLElement;
    grid.dispatchEvent(new MouseEvent('dblclick', { button: 0, bubbles: true }));
    const first = sheet.host.querySelector('.fc-host__editor');
    expect(first).not.toBeNull();

    // Second dblclick should be a no-op — the count stays at 1.
    grid.dispatchEvent(new MouseEvent('dblclick', { button: 0, bubbles: true }));
    const editors = sheet.host.querySelectorAll('.fc-host__editor');
    expect(editors.length).toBe(1);
  });
});

describe('mount/engine-binding — dispose unwires every subscription', () => {
  it('after dispose, wb mutations no longer reach the store', async () => {
    const sheet = await mountStubSheet();
    const wb = sheet.workbook;
    const store = sheet.instance.store;

    sheet.instance.dispose();

    // After dispose, the store is dead from the host's perspective. We can
    // still call wb.setNumber, but assertAlive() inside engine-binding's
    // subscriber would have already unsubscribed. The check here is that no
    // throw escapes — the regression that motivated dispose-leak.test.ts.
    expect(() => {
      wb.setNumber({ sheet: 0, row: 5, col: 5 }, 1);
    }).not.toThrow();

    // Store should be untouched by the post-dispose wb update.
    const cells = store.getState().data.cells;
    expect(cells.has('0:5:5')).toBe(false);
    sheet.dispose();
  });
});
