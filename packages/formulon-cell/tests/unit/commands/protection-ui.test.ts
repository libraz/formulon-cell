import { afterEach, describe, expect, it, vi } from 'vitest';
import { canExecuteBuiltIn } from '../../../src/commands/built-in-command-policy.js';
import {
  addThreadedComment,
  commentAt,
  recordThreadedChange,
  removeThreadedComment,
  setComment,
  threadedCommentsAt,
} from '../../../src/commands/comment.js';
import { commentWrite, threadWrite } from '../../../src/commands/comment-actions.js';
import { filterUseAllowed } from '../../../src/commands/filter-actions.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import type { InteractionPolicy } from '../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { setProtectionAlertSink } from '../../../src/commands/protection-alert.js';
import { executeRibbonClearAction } from '../../../src/commands/ribbon-clear.js';
import { setSlicerSelectionAction, slicerInteractive } from '../../../src/commands/slicers.js';
import { refreshAllConnections } from '../../../src/commands/table-refresh.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const A1 = { sheet: 0, row: 0, col: 0 };
const TABLE = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 };
const allowAll: InteractionPolicy = { defaultOperation: 'allow' };

describe('measured protection rules of the UI', () => {
  const cleanups: (() => void)[] = [];

  afterEach(() => {
    for (const cleanup of cleanups.splice(0)) cleanup();
  });

  const setup = async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    const controller = new InteractionController({ store, getWb: () => wb, history });
    registerInteractionController(store, controller);
    cleanups.push(() => {
      setProtectionAlertSink(wb, null);
      controller.dispose();
      wb.dispose();
    });
    return { wb, store, history, controller };
  };

  const withTable = async () => {
    const env = await setup();
    const { wb, store } = env;
    for (const [row, region] of ['Region', 'East', 'West', 'East'].entries())
      wb.setText({ sheet: 0, row, col: 0 }, region);
    expect(wb.createTable(0, { range: 'A1:B4', name: 'Sales', columns: ['Region', 'Qty'] })).toBe(
      0,
    );
    store.setState((s) => ({
      ...s,
      tables: {
        ...s.tables,
        tables: [
          ...s.tables.tables,
          {
            id: 'Sales',
            name: 'Sales',
            range: TABLE,
            source: 'engine',
          } as (typeof s.tables.tables)[number],
        ],
      },
    }));
    return env;
  };

  it('greys Clear Contents when any selected cell is locked: the clear is all or nothing', async () => {
    const { store } = await setup();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });
    setProtectedSheet(store, 0, true);
    mutators.setCellFormat(store, { sheet: 0, row: 1, col: 0 }, { locked: false });
    expect(canExecuteBuiltIn(store, 'clear', 'contextMenu').allowed).toBe(false);
    // The Home Clear control is greyed as a whole on the same selection.
    expect(canExecuteBuiltIn(store, 'clearFormat', 'ribbon').allowed).toBe(false);
    mutators.setRange(store, { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 });
    expect(canExecuteBuiltIn(store, 'clear', 'contextMenu').allowed).toBe(true);
    expect(canExecuteBuiltIn(store, 'clearFormat', 'ribbon').allowed).toBe(true);
  });

  it('opens no filter dropdown on a protected sheet without Use AutoFilter', async () => {
    const { store } = await setup();
    const entry = { store, origin: 'editor', commandId: 'filterDropdown' } as const;
    expect(filterUseAllowed(entry, TABLE)).toBe(true);
    setProtectedSheet(store, 0, true);
    expect(filterUseAllowed(entry, TABLE)).toBe(false);
    setProtectedSheet(store, 0, true, { permissions: { autoFilter: true } });
    expect(filterUseAllowed(entry, TABLE)).toBe(true);
  });

  for (const [label, policy] of [
    ['without a policy', undefined],
    ['under an allow-all policy', allowAll],
  ] as const satisfies readonly (readonly [string, InteractionPolicy | undefined])[]) {
    it(`lets threads through a protected sheet's locked cell, notes not, ${label}`, async () => {
      const { wb, store, history, controller } = await setup();
      controller.setPolicy(policy);
      setProtectedSheet(store, 0, true);
      const ctx = { store, history, origin: 'ribbon', commandId: 'mac.review.newComment' } as const;
      expect(commentWrite(ctx, [A1], false, () => true)).toBe(false);
      const added = threadWrite(ctx, [A1], null, () =>
        recordThreadedChange(history, store, wb, 0, () =>
          addThreadedComment(store, A1, 'hello', wb, { authorization: ctx }),
        ),
      );
      expect(added).not.toBeNull();
      expect(threadedCommentsAt(store, A1, wb)).toHaveLength(1);
      expect(history.undo()).toBe(true);
      expect(threadedCommentsAt(store, A1, wb)).toHaveLength(0);
      expect(history.redo()).toBe(true);
      const opening = threadedCommentsAt(store, A1, wb)[0];
      if (!opening) throw new Error('missing thread');
      expect(removeThreadedComment(store, 0, opening.id, wb, ctx)).toBe(true);
      // The projection of a thread command ignores protection too.
      mutators.setActive(store, A1);
      expect(canExecuteBuiltIn(store, 'mac.review.newComment', 'ribbon').allowed).toBe(true);
    });
  }

  it('still lets a policy refuse a thread', async () => {
    const { wb, store, history, controller } = await setup();
    controller.setPolicy({ defaultOperation: 'allow', operations: { comment: false } });
    const ctx = { store, history, origin: 'ribbon', commandId: 'mac.review.newComment' } as const;
    expect(
      threadWrite(ctx, [A1], null, () =>
        addThreadedComment(store, A1, 'hello', wb, { authorization: ctx }),
      ),
    ).toBeNull();
    expect(threadedCommentsAt(store, A1, wb)).toHaveLength(0);
  });

  it('greys Insert Slicer and Remove Duplicates over a table on a protected sheet', async () => {
    const { store } = await withTable();
    setProtectedSheet(store, 0, true, { permissions: { objects: true } });
    mutators.setActive(store, { sheet: 0, row: 1, col: 0 });
    for (const id of ['mac.insert.slicer', 'mac.data.removeDuplicates'] as const)
      expect(canExecuteBuiltIn(store, id, 'ribbon').allowed, id).toBe(false);
    expect(canExecuteBuiltIn(store, 'tableRefresh', 'ribbon').allowed).toBe(true);
    // Outside the table the slicer insert follows Edit objects as before.
    mutators.setActive(store, { sheet: 0, row: 10, col: 5 });
    expect(canExecuteBuiltIn(store, 'mac.insert.slicer', 'ribbon').allowed).toBe(true);
  });

  it('keeps a locked slicer inert without Edit objects and alerts when only the filter is refused', async () => {
    const { wb, store, history } = await withTable();
    mutators.addSlicer(store, {
      id: 's1',
      tableName: 'Sales',
      column: 'Region',
      selected: [],
      sheet: 0,
      x: 0,
      y: 0,
    });
    const alerts = vi.fn();
    setProtectionAlertSink(wb, alerts, { filterRefused: 'filter refused' });
    const ctx = {
      store,
      history,
      workbook: wb,
      origin: 'editor',
      commandId: 'slicerFilter',
    } as const;
    const spec = () => store.getState().slicers.slicers[0];
    const slicer = spec();
    if (!slicer) throw new Error('missing slicer');

    setProtectedSheet(store, 0, true);
    expect(slicerInteractive(ctx, slicer)).toBe(false);
    expect(setSlicerSelectionAction(ctx, 's1', ['West'])).toBe(false);
    expect(alerts).not.toHaveBeenCalled();

    setProtectedSheet(store, 0, true, { permissions: { objects: true } });
    expect(slicerInteractive(ctx, slicer)).toBe(true);
    expect(setSlicerSelectionAction(ctx, 's1', ['West'])).toBe(false);
    expect(alerts).toHaveBeenCalledWith({ severity: 'stop', message: 'filter refused' });
    expect(spec()?.selected).toEqual([]);

    setProtectedSheet(store, 0, true, { permissions: { objects: true, autoFilter: true } });
    expect(setSlicerSelectionAction(ctx, 's1', ['West'])).toBe(true);
    expect(spec()?.selected).toEqual(['West']);
  });

  it('leaves a workbook without data connections untouched on Refresh All', async () => {
    const { wb, store, history } = await withTable();
    const changes = vi.fn();
    const unsubscribe = wb.subscribe(changes);
    const before = store.getState();
    expect(refreshAllConnections(wb)).toBe(0);
    unsubscribe();
    expect(changes).not.toHaveBeenCalled();
    expect(store.getState()).toBe(before);
    expect(history.canUndo()).toBe(false);
  });

  it('greys note commands without Edit objects, keeps Show/Hide, and blocks a thread over a note', async () => {
    const { store } = await setup();
    const B2 = { sheet: 0, row: 1, col: 1 };
    mutators.setActive(store, A1);
    // A cell holding a note takes no thread, protected or not.
    mutators.setCellFormat(store, A1, { comment: 'memo' });
    expect(canExecuteBuiltIn(store, 'mac.insert.newComment', 'ribbon').allowed).toBe(false);
    mutators.setActive(store, B2);
    expect(canExecuteBuiltIn(store, 'mac.insert.newComment', 'ribbon').allowed).toBe(true);

    setProtectedSheet(store, 0, true);
    mutators.setActive(store, A1);
    for (const id of [
      'insertComment',
      'deleteComment',
      'mac.review.newNote',
      'deleteCommentReview',
    ] as const)
      expect(canExecuteBuiltIn(store, id, 'contextMenu').allowed, id).toBe(false);
    for (const id of ['showHideNote', 'mac.review.showHideNote', 'mac.review.showNotes'] as const)
      expect(canExecuteBuiltIn(store, id, 'contextMenu').allowed, id).toBe(true);
    // Review > Delete over a cell without a note is left to the threads it deletes.
    mutators.setActive(store, B2);
    expect(canExecuteBuiltIn(store, 'deleteCommentReview', 'ribbon').allowed).toBe(true);

    setProtectedSheet(store, 0, true, { permissions: { objects: true } });
    mutators.setActive(store, A1);
    for (const id of ['insertComment', 'deleteComment', 'deleteCommentReview'] as const)
      expect(canExecuteBuiltIn(store, id, 'contextMenu').allowed, id).toBe(true);
  });

  it('clears threads with Clear Comments in every mode while a locked note stays', async () => {
    const { wb, store, history } = await setup();
    const B1 = { sheet: 0, row: 0, col: 1 };
    setComment(store, A1, 'memo', wb);
    addThreadedComment(store, B1, 'thread', wb);
    setProtectedSheet(store, 0, true);
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    executeRibbonClearAction({ store, workbook: wb, history, action: 'comments' });
    expect(threadedCommentsAt(store, B1, wb)).toHaveLength(0);
    expect(commentAt(store.getState(), A1)).toBe('memo');
  });
});
