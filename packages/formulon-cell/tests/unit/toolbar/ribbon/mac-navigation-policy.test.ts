import { afterEach, describe, expect, it } from 'vitest';
import { noteVisibleAt, setComment } from '../../../../src/commands/comment.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createI18nController } from '../../../../src/i18n/controller.js';
import { attachNavigationPolicy } from '../../../../src/interact/navigation-policy.js';
import { createDefaultRibbonHooks } from '../../../../src/mount/toolbar-defaults.js';
import type { SpreadsheetInstance } from '../../../../src/mount/types.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';
import { cancelOpenAppDialogs } from '../../../../src/toolbar/dialogs/shell.js';
import {
  type ApplyRibbonCommandDeps,
  applyRibbonCommand,
} from '../../../../src/toolbar/ribbon/apply-ribbon-command.js';

afterEach(() => {
  cancelOpenAppDialogs();
  document.body.replaceChildren();
});

describe('Mac note navigation bounds', () => {
  it('skips excluded notes and opens only on an allowed note', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const host = document.createElement('div');
    document.body.append(host);
    const opened: string[] = [];
    const instance = {
      store,
      workbook,
      host,
      i18n: createI18nController({ locale: 'en' }),
      openCommentDialog: () => {
        const a = store.getState().selection.active;
        opened.push(`${a.row}:${a.col}`);
      },
    } as unknown as SpreadsheetInstance;
    const deps = {
      inst: instance,
      runtime: { projectFormatToolbar: () => {} },
      hooks: createDefaultRibbonHooks(instance),
    } as unknown as ApplyRibbonCommandDeps;
    setComment(store, { sheet: 0, row: 0, col: 2 }, 'excluded note');
    const navigation = attachNavigationPolicy(store, () => workbook, {
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
    });
    try {
      expect(applyRibbonCommand('mac.review.nextNote', deps)).toBe(true);
      expect(opened).toEqual([]);
      expect(store.getState().selection.active).toEqual({ sheet: 0, row: 0, col: 0 });
      setComment(store, { sheet: 0, row: 0, col: 1 }, 'allowed note');
      applyRibbonCommand('mac.review.nextNote', deps);
      expect(opened).toEqual(['0:1']);
      expect(store.getState().selection.active).toEqual({ sheet: 0, row: 0, col: 1 });
      applyRibbonCommand('mac.review.showNotes', deps);
      expect(noteVisibleAt(store.getState(), { sheet: 0, row: 0, col: 1 })).toBe(true);
    } finally {
      navigation.dispose();
      workbook.dispose();
    }
  });

  it('routes Mac note navigation through the shared review hook so a host override applies to both surfaces', () => {
    const calls: number[] = [];
    const deps = {
      inst: { store: createSpreadsheetStore() } as unknown as SpreadsheetInstance,
      runtime: { projectFormatToolbar: () => {} },
      hooks: { review: { selectComment: (direction: 1 | -1) => calls.push(direction) } },
    } as unknown as ApplyRibbonCommandDeps;
    expect(applyRibbonCommand('mac.review.nextNote', deps)).toBe(true);
    expect(applyRibbonCommand('mac.review.previousNote', deps)).toBe(true);
    expect(applyRibbonCommand('nextCommentReview', deps)).toBe(true);
    expect(calls).toEqual([1, -1, 1]);
  });

  it('moves between notes from the default shared ribbon hook', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const opened: string[] = [];
    const instance = {
      store,
      workbook,
      host: document.createElement('div'),
      i18n: createI18nController({ locale: 'en' }),
      openCommentDialog: () => {
        const a = store.getState().selection.active;
        opened.push(`${a.row}:${a.col}`);
      },
    } as unknown as SpreadsheetInstance;
    const deps = {
      inst: instance,
      runtime: { projectFormatToolbar: () => {} },
      hooks: createDefaultRibbonHooks(instance),
    } as unknown as ApplyRibbonCommandDeps;
    setComment(store, { sheet: 0, row: 0, col: 1 }, 'first');
    setComment(store, { sheet: 0, row: 3, col: 0 }, 'second');
    try {
      applyRibbonCommand('nextCommentReview', deps);
      applyRibbonCommand('nextCommentReview', deps);
      applyRibbonCommand('previousCommentReview', deps);
      expect(opened).toEqual(['0:1', '3:0', '0:1']);
    } finally {
      workbook.dispose();
    }
  });

  it('Show All Notes shows every note, and Show/Hide Note toggles the active one', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const host = document.createElement('div');
    document.body.append(host);
    const instance = {
      store,
      workbook,
      host,
      i18n: createI18nController({ locale: 'en-US' }),
    } as unknown as SpreadsheetInstance;
    const deps = {
      inst: instance,
      runtime: { projectFormatToolbar: () => {} },
    } as unknown as ApplyRibbonCommandDeps;
    setComment(store, { sheet: 0, row: 2, col: 1 }, 'note text');
    try {
      const note = { sheet: 0, row: 2, col: 1 };
      applyRibbonCommand('mac.review.showNotes', deps);
      expect(noteVisibleAt(store.getState(), note)).toBe(true);
      store.setState((s) => ({
        ...s,
        selection: { ...s.selection, active: note, anchor: note },
      }));
      applyRibbonCommand('mac.review.showHideNote', deps);
      expect(noteVisibleAt(store.getState(), note)).toBe(false);
    } finally {
      workbook.dispose();
    }
  });
});
