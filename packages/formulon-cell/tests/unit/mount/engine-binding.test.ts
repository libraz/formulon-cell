import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createSessionShape } from '../../../src/commands/session-illustration.js';
import { type ChangeEvent, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
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

  it('opens the Insert Copied Cells dialog from the public instance API', () => {
    sheet.instance.openInsertCopiedCells();
    const dialog = document.querySelector<HTMLElement>('.fc-insertcopied');
    expect(dialog).not.toBeNull();
    expect(dialog?.getAttribute('role')).toBe('dialog');
  });

  it('hands the shared history to the context menu so its edits undo', () => {
    const wb = sheet.workbook;
    const store = sheet.instance.store;
    const addr = { sheet: 0, row: 0, col: 0 };
    wb.setText(addr, 'a');
    wb.setText({ sheet: 0, row: 0, col: 1 }, 'b');
    mutators.setActive(store, addr);

    sheet.host.dispatchEvent(
      new MouseEvent('contextmenu', { clientX: 10, clientY: 30, bubbles: true, cancelable: true }),
    );
    document
      .querySelector<HTMLButtonElement>('.fc-ctxmenu__item[data-fc-action="rowInsertAbove"]')
      ?.click();
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'text', value: 'a' });

    sheet.instance.undo();
    expect(wb.getValue(addr)).toEqual({ kind: 'text', value: 'a' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 }).kind).toBe('blank');
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

  it('emits `recalc` with the edited cell once the pass has run', () => {
    const onRecalc = vi.fn();
    sheet.instance.on('recalc', onRecalc);
    sheet.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 5);
    expect(onRecalc).toHaveBeenCalledTimes(1);
    const dirty = onRecalc.mock.calls[0]?.[0]?.dirty as ReadonlySet<string>;
    expect([...dirty]).toEqual(['0:0:0']);
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

describe('mount/engine-binding — atomic value batches', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    sheet = await mountStubSheet({ workbook });
  });

  afterEach(() => sheet.dispose());

  it('publishes a mounted atomic batch once before its callbacks', () => {
    const formulaCalls = vi.spyOn(sheet.workbook, 'cellFormula');
    const bulkFormulaCalls = vi.spyOn(sheet.workbook, 'cellFormulas');
    const setState = vi.spyOn(sheet.instance.store, 'setState');
    const callbacks: ReadonlyMap<string, { value: unknown }>[] = [];
    const changes: unknown[] = [];
    sheet.instance.on('cellChange', (event) => {
      changes.push(event);
      callbacks.push(new Map(sheet.instance.store.getState().data.cells));
    });

    sheet.workbook.applyCellPatchAtomic([
      { addr: { sheet: 0, row: 0, col: 0 }, value: { kind: 'number', value: 1 }, formula: null },
      { addr: { sheet: 0, row: 1, col: 0 }, value: { kind: 'number', value: 2 }, formula: null },
    ]);

    const cells = sheet.instance.store.getState().data.cells;
    expect(changes).toHaveLength(2);
    expect(callbacks).toHaveLength(2);
    for (const snapshot of callbacks) {
      expect(snapshot.get('0:0:0')?.value).toEqual({ kind: 'number', value: 1 });
      expect(snapshot.get('0:1:0')?.value).toEqual({ kind: 'number', value: 2 });
    }
    expect(cells.get('0:0:0')?.value).toEqual({ kind: 'number', value: 1 });
    expect(cells.get('0:1:0')?.value).toEqual({ kind: 'number', value: 2 });
    expect(setState).toHaveBeenCalledTimes(1);
    expect(formulaCalls).toHaveBeenCalledTimes(0);
    // The mounted binding reads once; atomic before/after snapshots use strict readers.
    expect(bulkFormulaCalls).toHaveBeenCalledTimes(1);
  });

  it('rereads the outer batch after a reentrant ordinary write', async () => {
    sheet.dispose();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    let nested = false;
    workbook.subscribe((event) => {
      if (
        event.kind === 'value' &&
        event.atomicBatch?.index === 1 &&
        event.addr.row === 1 &&
        !nested
      ) {
        nested = true;
        workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 9);
      }
    });
    sheet = await mountStubSheet({ workbook });
    const setState = vi.spyOn(sheet.instance.store, 'setState');
    const changes: Array<{ value: unknown }> = [];
    sheet.instance.on('cellChange', (event) => changes.push({ value: event.value }));

    workbook.applyCellPatchAtomic([
      { addr: { sheet: 0, row: 0, col: 0 }, value: { kind: 'number', value: 1 }, formula: null },
      { addr: { sheet: 0, row: 1, col: 0 }, value: { kind: 'number', value: 2 }, formula: null },
    ]);

    expect(workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 9 });
    expect(workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'number', value: 2 });
    expect(sheet.instance.store.getState().data.cells.get('0:0:0')?.value).toEqual({
      kind: 'number',
      value: 9,
    });
    expect(sheet.instance.store.getState().data.cells.get('0:1:0')?.value).toEqual({
      kind: 'number',
      value: 2,
    });
    expect(changes.map((event) => event.value)).toEqual([
      { kind: 'number', value: 9 },
      { kind: 'number', value: 1 },
      { kind: 'number', value: 2 },
    ]);
    expect(setState).toHaveBeenCalledTimes(2);
  });

  it('keeps nested atomic overlap authoritative for both completed batches', async () => {
    sheet.dispose();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    let nested = false;
    workbook.subscribe((event) => {
      if (
        event.kind === 'value' &&
        event.atomicBatch?.index === 1 &&
        event.addr.row === 1 &&
        !nested
      ) {
        nested = true;
        workbook.applyCellPatchAtomic([
          {
            addr: { sheet: 0, row: 0, col: 0 },
            value: { kind: 'number', value: 9 },
            formula: null,
          },
          {
            addr: { sheet: 0, row: 2, col: 0 },
            value: { kind: 'number', value: 3 },
            formula: null,
          },
        ]);
      }
    });
    sheet = await mountStubSheet({ workbook });
    const setState = vi.spyOn(sheet.instance.store, 'setState');
    const changes: number[] = [];
    sheet.instance.on('cellChange', (event) => {
      if (event.value.kind === 'number') changes.push(event.value.value);
    });

    workbook.applyCellPatchAtomic([
      { addr: { sheet: 0, row: 0, col: 0 }, value: { kind: 'number', value: 1 }, formula: null },
      { addr: { sheet: 0, row: 1, col: 0 }, value: { kind: 'number', value: 2 }, formula: null },
    ]);

    expect(changes).toEqual([9, 3, 1, 2]);
    expect(sheet.instance.store.getState().data.cells.get('0:0:0')?.value).toEqual({
      kind: 'number',
      value: 9,
    });
    expect(sheet.instance.store.getState().data.cells.get('0:1:0')?.value).toEqual({
      kind: 'number',
      value: 2,
    });
    expect(sheet.instance.store.getState().data.cells.get('0:2:0')?.value).toEqual({
      kind: 'number',
      value: 3,
    });
    expect(setState).toHaveBeenCalledTimes(2);
  });

  it('rereads an outer address when a nested write arrives before its event', async () => {
    sheet.dispose();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    let nested = false;
    workbook.subscribe((event) => {
      if (event.kind === 'value' && event.atomicBatch?.index === 0 && !nested) {
        nested = true;
        workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 9);
      }
    });
    sheet = await mountStubSheet({ workbook });
    const changes: number[] = [];
    sheet.instance.on('cellChange', (event) => {
      if (event.value.kind === 'number') changes.push(event.value.value);
    });

    workbook.applyCellPatchAtomic([
      { addr: { sheet: 0, row: 0, col: 0 }, value: { kind: 'number', value: 1 }, formula: null },
      { addr: { sheet: 0, row: 1, col: 0 }, value: { kind: 'number', value: 2 }, formula: null },
    ]);

    expect(changes).toEqual([9, 1, 2]);
    expect(sheet.instance.store.getState().data.cells.get('0:1:0')?.value).toEqual({
      kind: 'number',
      value: 9,
    });
  });

  it('drops incomplete tagged batches on unbind without a deferred flush', async () => {
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    let listener: ((event: ChangeEvent) => void) | undefined;
    const subscribe = vi.spyOn(workbook, 'subscribe');
    subscribe.mockImplementation((callback) => {
      listener = callback;
      return () => {};
    });
    sheet.dispose();
    sheet = await mountStubSheet({ workbook });
    const setState = vi.spyOn(sheet.instance.store, 'setState');
    const first: ChangeEvent = {
      kind: 'value',
      addr: { sheet: 0, row: 0, col: 0 },
      next: { kind: 'number', value: 1 },
      atomicBatch: { id: 100, index: 0, size: 2, formula: null },
    };
    listener?.(first);
    expect(setState).toHaveBeenCalledTimes(0);
    sheet.dispose();
    listener?.({
      kind: 'value',
      addr: { sheet: 0, row: 1, col: 0 },
      next: { kind: 'number', value: 2 },
      atomicBatch: { id: 100, index: 1, size: 2, formula: null },
    });
    expect(setState).toHaveBeenCalledTimes(0);
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

  it('dblclick hides formula text when the active cell is Hidden on a protected sheet', () => {
    const addr = { sheet: 0, row: 0, col: 0 };
    sheet.workbook.setFormula(addr, '=2+3');
    sheet.workbook.recalc();
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setCellFormat(sheet.instance.store, addr, { formulaHidden: true });
    mutators.setSheetProtected(sheet.instance.store, 0, true);
    mutators.setActive(sheet.instance.store, addr);

    const grid = sheet.host.querySelector('.fc-host__grid') as HTMLElement;
    grid.dispatchEvent(new MouseEvent('dblclick', { button: 0, bubbles: true }));

    const editor = sheet.host.querySelector('.fc-host__editor') as HTMLTextAreaElement | null;
    expect(editor).not.toBeNull();
    expect(editor?.value).toBe('');
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

describe('mount/engine-binding — workbook replacement clears session objects', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
  });

  afterEach(() => sheet.dispose());

  it('drops session illustrations when a new workbook is bound', async () => {
    createSessionShape(
      sheet.instance.store,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      { id: 'shape-a', shape: 'rounded-rectangle', color: '#0f6cbd', radius: 12 },
      sheet.instance.history,
    );
    expect(sheet.instance.store.getState().illustrations.illustrations).toHaveLength(1);

    const next = await WorkbookHandle.createDefault({ preferStub: true });
    await sheet.instance.setWorkbook(next);

    expect(sheet.instance.store.getState().illustrations.illustrations).toEqual([]);
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
