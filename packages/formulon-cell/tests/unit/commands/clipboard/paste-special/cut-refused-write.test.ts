import { FormulonError, StatusCode } from '@libraz/formulon';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { pasteSpecial } from '../../../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../../../src/commands/clipboard/snapshot.js';
import { canUndo, History } from '../../../../../src/commands/history.js';
import { WorkbookHandle } from '../../../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../../src/store/store.js';
import { assertSnap, seedAndMirror, setActive } from './fixtures.js';

const ALL = { what: 'all', operation: 'none', skipBlanks: false, transpose: false } as const;

type EngineSheet = { setFormula(row: number, col: number, formula: string): void };
type SheetAccess = { ws(index: number): EngineSheet };

/** Make the engine refuse any formula written to `row`/`col` of sheet 0. */
const refuseFormulaAt = (wb: WorkbookHandle, row: number, col: number): void => {
  const access = wb as unknown as SheetAccess;
  const original = access.ws.bind(wb);
  vi.spyOn(access, 'ws').mockImplementation((index: number) => {
    const sheet = original(index);
    if (index !== 0) return sheet;
    return new Proxy(sheet, {
      get(target, prop) {
        const member = Reflect.get(target, prop, target);
        if (prop === 'setFormula') {
          return (r: number, c: number, formula: string) => {
            if (r === row && c === col) {
              throw new FormulonError(StatusCode.InvalidArgument, 'refused', '');
            }
            return member.call(target, r, c, formula);
          };
        }
        return typeof member === 'function' ? member.bind(target) : member;
      },
    });
  });
};

describe('cut paste refused by the engine mid-write', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault();
    if (wb.isStub) throw new Error('native workbook required');
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 1, col: 0, value: 2, formula: '=A1*2' },
      { row: 2, col: 0, value: 3, formula: '=A1*3' },
      { row: 0, col: 1, value: 5 },
      { row: 2, col: 1, value: 'x' },
      { row: 9, col: 3, value: 4, formula: '=A3+1' },
    ]);
    expect(wb.addSheet('Other')).toBe(1);
    wb.setFormula({ sheet: 1, row: 0, col: 0 }, `=${wb.sheetName(0)}!A2`);
    mutators.setCellFormat(store, { sheet: 0, row: 1, col: 0 }, { bold: true });
    const merge = { sheet: 0, r0: 0, c0: 1, r1: 1, c1: 1 };
    mutators.mergeRange(store, merge);
    wb.engineAddMerge(0, merge);
    wb.recalc();
  });

  afterEach(() => {
    vi.restoreAllMocks();
    wb.dispose();
  });

  const sheetContent = () => {
    const out: Record<string, unknown> = {};
    for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) {
      for (let row = 0; row < 12; row += 1) {
        for (let col = 0; col < 8; col += 1) {
          const addr = { sheet, row, col };
          out[`${sheet}:${row}:${col}`] = [wb.cellFormula(addr), wb.getValue(addr)];
        }
      }
    }
    return out;
  };
  const sideState = () => {
    const s = store.getState();
    return {
      formats: [...s.format.formats],
      merges: [...s.merges.byAnchor],
      engineMerges: wb.getMerges(0),
    };
  };

  /** Cut A1:B3 to E1; the second destination formula (E3) is refused. */
  const cutRefused = (history: History | null): void => {
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 }, 'cut');
    assertSnap(snap);
    setActive(store, 0, 4);
    refuseFormulaAt(wb, 2, 4);
    expect(() => pasteSpecial(store.getState(), store, wb, snap, ALL, history)).toThrow('refused');
  };

  it('leaves the source, destination and other sheets untouched without history', () => {
    const cells = sheetContent();
    const side = sideState();

    cutRefused(null);

    expect(sheetContent()).toEqual(cells);
    expect(sideState()).toEqual(side);
  });

  it('leaves the sheet and the history untouched under a history', () => {
    const history = new History();
    wb.attachHistory(history);
    const cells = sheetContent();
    const side = sideState();

    cutRefused(history);

    expect(sheetContent()).toEqual(cells);
    expect(sideState()).toEqual(side);
    expect(canUndo(history)).toBe(false);
  });

  it('moves the block, its merge and the references to it as one undo step', () => {
    const history = new History();
    wb.attachHistory(history);
    const cells = sheetContent();
    const side = sideState();
    const snap = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 }, 'cut');
    assertSnap(snap);
    setActive(store, 0, 4);

    history.begin();
    const result = pasteSpecial(store.getState(), store, wb, snap, ALL, history);
    history.end();

    expect(result.range).toEqual({ sheet: 0, r0: 0, c0: 4, r1: 2, c1: 5 });
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 4 })).toBe('=E1*3');
    expect(wb.getValue({ sheet: 0, row: 2, col: 4 })).toEqual({ kind: 'number', value: 3 });
    expect(wb.cellFormula({ sheet: 0, row: 9, col: 3 })).toBe('=E3+1');
    expect(wb.cellFormula({ sheet: 1, row: 0, col: 0 })).toBe(`=${wb.sheetName(0)}!E2`);
    expect(wb.getMerges(0)).toEqual([{ sheet: 0, r0: 0, c0: 5, r1: 1, c1: 5 }]);

    expect(history.undo()).toBe(true);
    expect(sheetContent()).toEqual(cells);
    expect(sideState()).toEqual(side);
    expect(canUndo(history)).toBe(false);
  });
});
