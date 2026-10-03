import { describe, expect, it } from 'vitest';
import {
  applySelectionFormatAction,
  planSelectionFormat,
} from '../../../../src/commands/format.js';
import { History } from '../../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../../src/commands/interaction-controller.js';
import { fixedFormPolicy } from '../../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../../src/store/store.js';
import { fmtAt, setSelection } from './fixtures.js';

describe('selection format actions', () => {
  it('does not stage a locked blank cell and only stages an unlocked blank cell', () => {
    const locked = createSpreadsheetStore();
    setSelection(locked, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setProtectedSheet(locked, 0, true);
    expect(applySelectionFormatAction(locked.getState(), locked, { patch: { bold: true } })).toBe(
      false,
    );
    expect(locked.getState().ui.pendingFormat).toBeNull();
    expect(locked.getState().format.formats.size).toBe(0);
    expect(
      applySelectionFormatAction(
        locked.getState(),
        locked,
        { patch: { bold: true } },
        { allowPending: false },
      ),
    ).toBe(false);
    expect(locked.getState().ui.pendingFormat).toBeNull();
    expect(locked.getState().format.formats.size).toBe(0);

    const unlocked = createSpreadsheetStore();
    setSelection(unlocked, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    expect(
      applySelectionFormatAction(unlocked.getState(), unlocked, { patch: { bold: true } }),
    ).toBe(true);
    expect(unlocked.getState().ui.pendingFormat?.format).toEqual({ bold: true });
  });

  it('plans the disjoint union and preserves untouched mixed metadata', () => {
    const store = createSpreadsheetStore();
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 }, [
      { sheet: 0, r0: 4, c0: 4, r1: 4, c1: 4 },
    ]);
    const active = { sheet: 0, row: 0, col: 0 };
    const extra = { sheet: 0, row: 4, col: 4 };
    mutators.setCellFormat(store, active, {
      bold: true,
      hyperlink: 'https://example.test',
      comment: 'keep active metadata',
    });
    mutators.setCellFormat(store, extra, {
      bold: false,
      hyperlink: 'https://extra.test',
      comment: 'keep extra metadata',
    });

    const plan = planSelectionFormat(store.getState());
    expect(plan?.cells).toHaveLength(10);
    expect(
      applySelectionFormatAction(store.getState(), store, {
        patch: { align: 'center' },
      }),
    ).toBe(true);
    expect(fmtAt(store, 0, 0)).toMatchObject({
      align: 'center',
      bold: true,
      hyperlink: 'https://example.test',
    });
    expect(fmtAt(store, 4, 4)).toMatchObject({
      align: 'center',
      bold: false,
      hyperlink: 'https://extra.test',
    });
    expect(fmtAt(store, 1, 1)?.align).toBe('center');
  });

  it('derives metadata permissions and cannot bypass them with optional operations', async () => {
    const store = createSpreadsheetStore();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const controller = new InteractionController({
      store,
      getWb: () => workbook,
      history: new History(),
    });
    const unregister = registerInteractionController(store, controller);
    try {
      setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
      controller.setPolicy({
        ...fixedFormPolicy([{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }]),
        operations: { format: true },
      });
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
      const before = new Map(store.getState().format.formats);

      expect(
        applySelectionFormatAction(store.getState(), store, {
          patch: { bold: false, comment: 'blocked' },
          operations: ['format'],
        }),
      ).toBe(false);
      expect(store.getState().format.formats).toEqual(before);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  it('preserves each target hyperlink metadata only when its URL is unchanged', () => {
    const store = createSpreadsheetStore();
    setSelection(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, [
      { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
      { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 },
    ]);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        hyperlink: 'https://same.example',
        hyperlinkDisplay: 'A',
        hyperlinkTooltip: 'tip A',
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 1 },
      {
        hyperlink: 'https://same.example',
        hyperlinkDisplay: 'B',
        hyperlinkTooltip: 'tip B',
      },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 2 },
      {
        hyperlink: 'https://old.example',
        hyperlinkDisplay: 'C',
        hyperlinkTooltip: 'tip C',
      },
    );

    expect(
      applySelectionFormatAction(store.getState(), store, {
        patch: { hyperlink: 'https://same.example' },
      }),
    ).toBe(true);
    expect(fmtAt(store, 0, 0)).toMatchObject({
      hyperlink: 'https://same.example',
      hyperlinkDisplay: 'A',
      hyperlinkTooltip: 'tip A',
    });
    expect(fmtAt(store, 0, 1)).toMatchObject({
      hyperlink: 'https://same.example',
      hyperlinkDisplay: 'B',
      hyperlinkTooltip: 'tip B',
    });
    expect(fmtAt(store, 0, 2)).toMatchObject({ hyperlink: 'https://same.example' });
    expect(fmtAt(store, 0, 2)?.hyperlinkDisplay).toBeUndefined();
    expect(fmtAt(store, 0, 2)?.hyperlinkTooltip).toBeUndefined();
  });
});
