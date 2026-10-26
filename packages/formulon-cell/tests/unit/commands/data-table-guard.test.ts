import { describe, expect, it } from 'vitest';
import { commitCellInput } from '../../../src/commands/cell-input-commit.js';
import { insertCells } from '../../../src/commands/cell-shift.js';
import { clearSelectedContents } from '../../../src/commands/clear-contents.js';
import { copy } from '../../../src/commands/clipboard/copy.js';
import { insertCopiedBand } from '../../../src/commands/clipboard/insert-copied-cells.js';
import { pasteTSV } from '../../../src/commands/clipboard/paste.js';
import { pasteSpecial } from '../../../src/commands/clipboard/paste-special.js';
import {
  captureSnapshot as captureBandSnapshot,
  captureSnapshot,
} from '../../../src/commands/clipboard/snapshot.js';
import { setDataTableGuardSink } from '../../../src/commands/data-table-guard.js';
import { fillRange } from '../../../src/commands/fill.js';
import { copyAdvancedFilterResult } from '../../../src/commands/filter.js';
import { replaceAll, replaceOne } from '../../../src/commands/find.js';
import { History } from '../../../src/commands/history.js';
import { importCSV } from '../../../src/commands/import-export.js';
import { InteractionController } from '../../../src/commands/interaction-controller.js';
import { planMacSubtotal } from '../../../src/commands/mac-data-tools.js';
import { applyMerge, applyMergeAcross } from '../../../src/commands/merge.js';
import { executeQuickAnalysisAction } from '../../../src/commands/quick-analysis.js';
import { sortRange } from '../../../src/commands/sort.js';
import { deleteCols, deleteRows, insertRows } from '../../../src/commands/structure.js';
import { applyTextScriptToRange } from '../../../src/commands/text-script.js';
import { textToColumns } from '../../../src/commands/text-to-columns.js';
import type { Addr } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en, ja } from '../../../src/i18n/strings.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';
const at = (row: number, col: number): Addr => ({ sheet: 0, row, col });

/** A1 = 5 input; B3:B5 = 1..3; C2 = A1*2; results in C3:C5. */
const setup = async () => {
  const wb = await WorkbookHandle.createDefault();
  wb.setNumber(at(0, 0), 5);
  for (const i of [1, 2, 3]) wb.setNumber(at(1 + i, 1), i);
  wb.setFormula(at(1, 2), '=A1*2');
  wb.addDataTable({
    range: { sheet: 0, r0: 2, c0: 2, r1: 4, c1: 2 },
    rowInput: null,
    colInput: at(0, 0),
  });
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  const alerts: string[] = [];
  setDataTableGuardSink(wb, (alert) => alerts.push(alert.message), en.dataTableDialog);
  return { wb, store, alerts };
};

describe.skipIf(!canLoadWasm())('data table edit guard', () => {
  it('rejects an editor commit onto a result cell', async () => {
    const { wb, store } = await setup();
    const result = commitCellInput({ store, wb, addr: at(3, 2), raw: '9', origin: 'editor' });
    expect(result).toMatchObject({
      status: 'rejected',
      alert: { severity: 'stop', message: en.dataTableDialog.cannotChange },
    });
    expect(wb.getValue(at(3, 2))).toEqual({ kind: 'number', value: 4 });
    // The formula row and the substitution values stay editable.
    expect(
      commitCellInput({ store, wb, addr: at(1, 2), raw: '=A1*3', origin: 'editor' }).status,
    ).toBe('applied');
    expect(wb.getValue(at(3, 2))).toEqual({ kind: 'number', value: 6 });
    wb.dispose();
  });

  it('words the alert in the UI language', async () => {
    const { wb, store } = await setup();
    setDataTableGuardSink(wb, () => undefined, ja.dataTableDialog);
    const result = commitCellInput({ store, wb, addr: at(2, 2), raw: '1', origin: 'formulaBar' });
    expect(result).toMatchObject({ alert: { message: ja.dataTableDialog.cannotChange } });
    wb.dispose();
  });

  it('refuses Delete over part of the results and removes a wholly selected table', async () => {
    const { wb, store, alerts } = await setup();
    const history = new History();
    wb.attachHistory(history);
    mutators.setRange(store, { sheet: 0, r0: 3, c0: 1, r1: 4, c1: 2 });
    expect(clearSelectedContents({ store, workbook: wb, history }).status).toBe('rejected');
    expect(alerts).toEqual([en.dataTableDialog.cannotChange]);
    expect(wb.getValue(at(4, 1))).toEqual({ kind: 'number', value: 3 });

    alerts.length = 0;
    mutators.setRange(store, { sheet: 0, r0: 1, c0: 1, r1: 4, c1: 2 });
    expect(clearSelectedContents({ store, workbook: wb, history }).status).toBe('applied');
    expect(alerts).toEqual([]);
    expect(wb.dataTables(0)).toEqual([]);
    expect(wb.getValue(at(3, 2))).toEqual({ kind: 'blank' });
    expect(wb.getValue(at(3, 1))).toEqual({ kind: 'blank' });

    history.undo();
    expect(wb.dataTables(0).map((t) => t.range)).toEqual([
      { sheet: 0, r0: 2, c0: 2, r1: 4, c1: 2 },
    ]);
    expect(wb.getValue(at(3, 2))).toEqual({ kind: 'number', value: 4 });
    history.redo();
    expect(wb.dataTables(0)).toEqual([]);
    wb.dispose();
  });

  it('refuses a paste that reaches the results', async () => {
    const { wb, store, alerts } = await setup();
    mutators.setActive(store, at(4, 1));
    expect(pasteTSV(store.getState(), wb, 'a\tb')).toBeNull();
    expect(alerts).toEqual([en.dataTableDialog.cannotChange]);
    expect(wb.getValue(at(4, 1))).toEqual({ kind: 'number', value: 3 });
    mutators.setActive(store, at(10, 0));
    expect(pasteTSV(store.getState(), wb, 'a')).not.toBeNull();
    wb.dispose();
  });

  it('refuses row and column edits that would split the table', async () => {
    const { wb, store, alerts } = await setup();
    expect(insertRows(store, wb, null, 3)).toBe(false);
    expect(deleteRows(store, wb, null, 1)).toBe(false);
    expect(deleteCols(store, wb, null, 1)).toBe(false);
    expect(alerts).toEqual([
      en.dataTableDialog.cannotChange,
      en.dataTableDialog.cannotChange,
      en.dataTableDialog.cannotChange,
    ]);
    expect(wb.dataTableAt(at(2, 2))?.range).toEqual({ sheet: 0, r0: 2, c0: 2, r1: 4, c1: 2 });

    alerts.length = 0;
    expect(insertRows(store, wb, null, 1)).toBe(true);
    expect(alerts).toEqual([]);
    expect(wb.dataTableAt(at(3, 2))?.range).toEqual({ sheet: 0, r0: 3, c0: 2, r1: 5, c1: 2 });
    wb.dispose();
  });

  it('refuses fills, cell shifts, sorts and cell batches that reach the results', async () => {
    const { wb, store, alerts } = await setup();
    const results = { sheet: 0, r0: 2, c0: 2, r1: 4, c1: 2 };
    // Filling the formula row down over the results.
    expect(
      fillRange(store.getState(), wb, { ...results, r0: 1, r1: 1 }, { ...results, r0: 1 }),
    ).toBe(false);
    // Shifting the input column down moves part of the table.
    expect(insertCells(store, wb, null, { sheet: 0, r0: 3, c0: 1, r1: 3, c1: 1 }, 'down')).toBe(
      false,
    );
    expect(
      sortRange(
        store.getState(),
        store,
        wb,
        { sheet: 0, r0: 2, c0: 1, r1: 4, c1: 2 },
        {
          byCol: 1,
          direction: 'desc',
        },
      ),
    ).toBe(false);
    const controller = new InteractionController({
      store,
      getWb: () => wb,
      history: new History(),
    });
    const batch = controller.execute({
      type: 'cellBatch',
      operation: 'valueEdit',
      origin: 'instanceApi',
      changes: [
        { addr: at(8, 8), input: '1' },
        { addr: at(4, 2), input: '2' },
      ],
    });
    expect(batch.status).toBe('rejected');
    expect(wb.getValue(at(8, 8))).toEqual({ kind: 'blank' });
    expect(alerts).toHaveLength(4);
    expect(wb.getValue(at(4, 2))).toEqual({ kind: 'number', value: 6 });
    // Shifting cells clear of the table still works.
    expect(insertCells(store, wb, null, { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 }, 'down')).toBe(
      true,
    );
    wb.dispose();
  });

  it('refuses a typed TABLE() formula the way the desktop app does', async () => {
    const { wb, store } = await setup();
    const result = commitCellInput({
      store,
      wb,
      addr: at(8, 8),
      raw: '=TABLE(,A1)',
      origin: 'editor',
    });
    expect(result).toMatchObject({
      status: 'rejected',
      alert: { message: en.dataTableDialog.tableFunction },
    });
    expect(wb.getValue(at(8, 8))).toEqual({ kind: 'blank' });
    const controller = new InteractionController({
      store,
      getWb: () => wb,
      history: new History(),
    });
    const batch = controller.execute({
      type: 'cellBatch',
      operation: 'formulaEdit',
      origin: 'instanceApi',
      changes: [{ addr: at(8, 8), input: '=table(A1,)' }],
    });
    expect(batch.status).toBe('rejected');
    expect(wb.getValue(at(8, 8))).toEqual({ kind: 'blank' });
    wb.dispose();
  });

  it('copies result cells as their values', async () => {
    const { wb, store } = await setup();
    mutators.setRange(store, { sheet: 0, r0: 2, c0: 2, r1: 4, c1: 2 });
    expect(copy(store.getState())?.tsv).toBe('2\r\n4\r\n6');
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 2, c0: 2, r1: 4, c1: 2 });
    if (!snap) throw new Error('no snapshot');
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 6, r1: 0, c1: 6 });
    mutators.setActive(store, at(0, 6));
    pasteSpecial(store.getState(), store, wb, snap, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    expect(wb.getValue(at(1, 6))).toEqual({ kind: 'number', value: 4 });
    expect(wb.dataTables(0)).toHaveLength(1);
    expect(wb.dataTableAt(at(1, 6))).toBeNull();
    wb.dispose();
  });

  it('refuses every other writer up front with the standard alert', async () => {
    const { wb, store, alerts } = await setup();
    const results = { sheet: 0, r0: 2, c0: 2, r1: 4, c1: 2 };
    const refresh = () => mutators.replaceCells(store, wb.cells(0));
    const expectRefused = (label: string, outcome: unknown, refused: unknown) => {
      expect(outcome, label).toEqual(refused);
      expect(alerts.length, label).toBe(1);
      alerts.length = 0;
    };

    expectRefused('merge', applyMerge(store, wb, null, { ...results, r0: 1, r1: 3 }), false);
    expectRefused(
      'merge across',
      applyMergeAcross(store, wb, null, { sheet: 0, r0: 3, c0: 1, r1: 3, c1: 2 }),
      false,
    );
    expectRefused(
      'replace',
      replaceOne(wb, { addr: at(3, 2) }, '=TABLE(,B1)', store, 'en-US', 'formulas'),
      false,
    );
    // Replace All touches the formula row too, but refuses before writing anything.
    expectRefused(
      'replace all',
      replaceAll(store.getState(), wb, { query: 'A1', lookIn: 'formulas' }, 'B1', store),
      0,
    );
    expect(wb.cellFormula(at(1, 2))).toBe('=A1*2');

    wb.setText(at(3, 1), 'a,b');
    refresh();
    expectRefused(
      'text to columns',
      textToColumns(store.getState(), store, wb, { sheet: 0, r0: 3, c0: 1, r1: 3, c1: 1 }, ','),
      0,
    );
    expect(wb.getValue(at(3, 1))).toEqual({ kind: 'text', value: 'a,b' });
    wb.setNumber(at(3, 1), 2);
    refresh();

    expectRefused(
      'quick analysis totals',
      executeQuickAnalysisAction({
        store,
        wb,
        actionId: 'totals-sum-col',
        range: { sheet: 0, r0: 2, c0: 0, r1: 4, c1: 1 },
        stats: { cells: 6, numericCount: 4, nonBlankCount: 4, sum: 0, avg: 0, min: 0, max: 0 },
      }),
      { ok: false, reason: 'protected' },
    );
    expectRefused('text script', applyTextScriptToRange(store.getState(), wb, results, 'clear'), 0);
    expectRefused('csv import', importCSV(store.getState(), wb, '1,2', at(4, 1)), null);
    expectRefused(
      'advanced filter copy',
      copyAdvancedFilterResult(
        store.getState(),
        store,
        { sheet: 0, r0: 1, c0: 1, r1: 4, c1: 1 },
        { sheet: 0, r0: 20, c0: 0, r1: 20, c1: 0 },
        at(2, 2),
        {},
        wb,
      ),
      0,
    );
    const instance = {
      store,
      workbook: wb,
      history: new History(),
    } as unknown as SpreadsheetInstance;
    const subtotal = planMacSubtotal(instance, {
      range: 'B2:C5',
      groupByColumn: 1,
      function: 'sum',
      subtotalColumns: [2],
    });
    expect(subtotal.ok).toBe(false);
    expect(alerts).toEqual([en.dataTableDialog.cannotChange]);
    alerts.length = 0;

    const history = new History();
    wb.attachHistory(history);
    const band = captureBandSnapshot(
      store.getState(),
      { sheet: 0, r0: 3, c0: 0, r1: 3, c1: 16383 },
      'cut',
    );
    expectRefused(
      'cut rows',
      insertCopiedBand(store, wb, history, band, { sheet: 0, row: 10, col: 0 }),
      null,
    );
    expect(wb.dataTableAt(at(3, 2))?.range).toEqual(results);
    wb.dispose();
  });

  it('finds TABLE() anywhere in a typed or pasted formula, but not inside text', async () => {
    const { wb, store, alerts } = await setup();
    for (const raw of ['=1+TABLE(,A1)', '=IF(TRUE, table (A1,))']) {
      const result = commitCellInput({ store, wb, addr: at(8, 8), raw, origin: 'editor' });
      expect(result, raw).toMatchObject({
        status: 'rejected',
        alert: { message: en.dataTableDialog.tableFunction },
      });
    }
    const text = commitCellInput({
      store,
      wb,
      addr: at(8, 8),
      raw: '="TABLE(" & MYTABLE',
      origin: 'editor',
    });
    expect(text.status).toBe('applied');

    mutators.setActive(store, at(10, 0));
    expect(pasteTSV(store.getState(), wb, '1\t=SUM(1,TABLE(,A1))')).toBeNull();
    expect(alerts).toEqual([en.dataTableDialog.tableFunction]);
    expect(wb.getValue(at(10, 0))).toEqual({ kind: 'blank' });
    wb.dispose();
  });
});
