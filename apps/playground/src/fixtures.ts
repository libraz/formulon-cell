/**
 * Deterministic seeds for visual-regression specs. Triggered by
 * `?fixture=<name>` in the URL. Each fixture writes a known shape into the
 * workbook + store so the canvas paints the same pixels across runs.
 */

import {
  addConditionalRule,
  mutators,
  type SpreadsheetInstance,
  type WorkbookHandle,
} from '@libraz/formulon-cell';

export type FixtureName = 'basic' | 'cf' | 'sparkline' | 'selection' | 'frozen';

export const isFixtureName = (s: string | null): s is FixtureName =>
  s === 'basic' || s === 'cf' || s === 'sparkline' || s === 'selection' || s === 'frozen';

/** Seed for V05 — a conditional-format active state.
 *  A 1×6 column where values > 10 paint in red. */
export function seedFixtureCf(wb: WorkbookHandle, inst: SpreadsheetInstance): void {
  wb.setText({ sheet: 0, row: 0, col: 0 }, 'value');
  for (let r = 0; r < 6; r += 1) {
    wb.setNumber({ sheet: 0, row: r + 1, col: 0 }, [3, 7, 12, 5, 15, 1][r] ?? 0);
  }
  wb.recalc();
  addConditionalRule(inst.store, {
    kind: 'cell-value',
    range: { sheet: 0, r0: 1, c0: 0, r1: 6, c1: 0 },
    op: '>',
    a: 10,
    apply: { fill: '#ffcccc', color: '#a00' },
  });
  mutators.replaceCells(inst.store, wb.cells(0));
}

/** Seed for V06 — a single sparkline cell.
 *  Source range B1:F1 holds 5 numbers; A1 hosts a line sparkline. */
export function seedFixtureSparkline(wb: WorkbookHandle, inst: SpreadsheetInstance): void {
  const data = [1, 3, 2, 5, 4];
  for (let i = 0; i < data.length; i += 1) {
    wb.setNumber({ sheet: 0, row: 0, col: i + 1 }, data[i] ?? 0);
  }
  wb.recalc();
  mutators.replaceCells(inst.store, wb.cells(0));
  mutators.setSparkline(
    inst.store,
    { sheet: 0, row: 0, col: 0 },
    { kind: 'line', source: 'B1:F1', color: '#0078d4' },
  );
}

/** Seed for V07 — a 3×3 range selected with the fill handle hint visible. */
export function seedFixtureSelection(wb: WorkbookHandle, inst: SpreadsheetInstance): void {
  for (let r = 0; r < 3; r += 1) {
    for (let c = 0; c < 3; c += 1) {
      wb.setNumber({ sheet: 0, row: r, col: c }, (r + 1) * (c + 1));
    }
  }
  wb.recalc();
  mutators.replaceCells(inst.store, wb.cells(0));
  mutators.setActive(inst.store, { sheet: 0, row: 0, col: 0 });
  mutators.setRange(inst.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 });
}

/** Seed for V08 — a frozen 2-row × 1-col pane. Frozen panes need engine
 *  capability support; under the stub the freeze call no-ops, so the visual
 *  baseline reflects "no freeze" then. The test still runs deterministically. */
export function seedFixtureFrozen(wb: WorkbookHandle, inst: SpreadsheetInstance): void {
  wb.setText({ sheet: 0, row: 0, col: 0 }, 'hdr');
  for (let r = 1; r < 8; r += 1) {
    wb.setNumber({ sheet: 0, row: r, col: 0 }, r);
  }
  wb.recalc();
  mutators.replaceCells(inst.store, wb.cells(0));
  // Best-effort: the engine may not support freeze under stub.
  wb.setSheetFreeze(0, 2, 1);
}

/** Apply the fixture identified by `name`. Returns true when a seed ran. */
export function applyFixture(
  name: FixtureName,
  wb: WorkbookHandle,
  inst: SpreadsheetInstance,
): boolean {
  switch (name) {
    case 'basic':
      // The basic fixture is whatever the default app seed already provides.
      return false;
    case 'cf':
      seedFixtureCf(wb, inst);
      return true;
    case 'sparkline':
      seedFixtureSparkline(wb, inst);
      return true;
    case 'selection':
      seedFixtureSelection(wb, inst);
      return true;
    case 'frozen':
      seedFixtureFrozen(wb, inst);
      return true;
    default:
      return false;
  }
}
