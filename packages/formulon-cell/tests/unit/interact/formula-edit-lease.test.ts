import { afterEach, describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createFormulaEditLease,
  type FormulaEditLeaseSnapshot,
} from '../../../src/interact/formula-edit-lease.js';

const addr = { sheet: 0, row: 2, col: 3 };

const snapshot = (workbook: WorkbookHandle): FormulaEditLeaseSnapshot => ({
  source: 'inline',
  workbook,
  anchor: { ...addr },
  raw: '=SUM(A1:B2)',
  baseline: '=SUM(',
  caret: { start: 6, end: 10, direction: 'backward' },
  selection: {
    active: { ...addr },
    anchor: { ...addr },
    range: { sheet: 0, r0: 2, c0: 3, r1: 4, c1: 5 },
    extraRanges: [{ sheet: 0, r0: 8, c0: 1, r1: 8, c1: 2 }],
  },
  editorMode: { kind: 'edit', raw: '=SUM(A1:B2)', caret: 10 },
  pendingFormat: { addr: { ...addr }, format: { bold: true } },
  editorRefs: [{ r0: 0, c0: 0, r1: 1, c1: 1, colorIndex: 2 }],
  copy: {
    copyRange: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 2 },
    copyRanges: [{ sheet: 0, r0: 5, c0: 0, r1: 5, c1: 0 }],
    copyMode: 'copy',
    copyRevision: 7,
  },
  r1c1: true,
  locale: 'ja',
});

describe('formula edit lease', () => {
  const workbooks: WorkbookHandle[] = [];

  afterEach(() => {
    for (const workbook of workbooks.splice(0)) workbook.dispose();
  });

  it('deep-clones owner data while retaining workbook identity and releases once', async () => {
    const workbook = await WorkbookHandle.createDefault();
    workbooks.push(workbook);
    const original = snapshot(workbook);
    const ownerCurrent = true;
    const restore = vi.fn(() => document.createElement('textarea'));
    const release = vi.fn();
    const lease = createFormulaEditLease(original, {
      isOwnerCurrent: () => ownerCurrent,
      context: { getLocale: () => 'ja', contextCurrent: () => true },
      restore,
      release,
    });

    (original.selection.active as { row: number }).row = 99;
    if (original.pendingFormat?.format) original.pendingFormat.format.bold = false;
    expect(lease.valid()).toBe(true);
    expect(lease.snapshot.workbook).toBe(workbook);
    expect(lease.snapshot.selection.active.row).toBe(2);
    expect(lease.snapshot.pendingFormat?.format.bold).toBe(true);

    lease.finalize();
    lease.finalize();
    expect(release).toHaveBeenCalledTimes(1);
    expect(lease.valid()).toBe(false);
    expect(lease.userCancel()).toBeNull();
    expect(restore).not.toHaveBeenCalled();
  });

  it('invalid context cancellation discards without restoring DOM or store state', async () => {
    const workbook = await WorkbookHandle.createDefault();
    workbooks.push(workbook);
    let contextCurrent = true;
    const restore = vi.fn(() => document.createElement('textarea'));
    const release = vi.fn();
    const lease = createFormulaEditLease(snapshot(workbook), {
      isOwnerCurrent: () => true,
      context: { getLocale: () => 'ja', contextCurrent: () => contextCurrent },
      restore,
      release,
    });

    expect(lease.valid()).toBe(true);
    contextCurrent = false;
    expect(lease.valid()).toBe(false);
    expect(lease.userCancel()).toBeNull();
    expect(restore).not.toHaveBeenCalled();
    expect(release).toHaveBeenCalledTimes(1);
    lease.discard();
    expect(release).toHaveBeenCalledTimes(1);
  });

  it('a locale switch after suspension invalidates the lease', async () => {
    const workbook = await WorkbookHandle.createDefault();
    workbooks.push(workbook);
    let locale = 'ja';
    const restore = vi.fn(() => document.createElement('textarea'));
    const release = vi.fn();
    const lease = createFormulaEditLease(snapshot(workbook), {
      isOwnerCurrent: () => true,
      context: { getLocale: () => locale, contextCurrent: () => true },
      restore,
      release,
    });

    expect(lease.valid()).toBe(true);
    locale = 'en';
    expect(lease.valid()).toBe(false);
    expect(lease.userCancel()).toBeNull();
    expect(restore).not.toHaveBeenCalled();
    expect(release).toHaveBeenCalledOnce();
  });

  it('a disposed workbook makes a suspended owner stale before user cancel', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const restore = vi.fn(() => document.createElement('textarea'));
    const lease = createFormulaEditLease(snapshot(workbook), {
      isOwnerCurrent: () => true,
      restore,
      release: () => {},
    });

    workbook.dispose();
    expect(lease.valid()).toBe(false);
    expect(lease.userCancel()).toBeNull();
    expect(restore).not.toHaveBeenCalled();
  });

  it('restores and releases only once when restoration synchronously cancels again', async () => {
    const workbook = await WorkbookHandle.createDefault();
    workbooks.push(workbook);
    const element = document.createElement('textarea');
    let cancelAgain = () => {};
    const restore = vi.fn(() => {
      cancelAgain();
      return element;
    });
    const release = vi.fn();
    const lease = createFormulaEditLease(snapshot(workbook), {
      isOwnerCurrent: () => true,
      restore,
      release,
    });
    cancelAgain = () => {
      expect(lease.userCancel()).toBeNull();
    };
    expect(lease.userCancel()).toBe(element);
    expect(restore).toHaveBeenCalledOnce();
    expect(release).toHaveBeenCalledOnce();
    expect(lease.userCancel()).toBeNull();
  });
});
