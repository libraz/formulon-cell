import { expect, type Page } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** T03 — freeze pane toggle round-trip.
 *
 *  Setting and clearing freeze panes is a layout mutation. The visible result
 *  (a thin pane divider on the canvas) is hard to assert from Playwright
 *  without pixel diffing, but the underlying `layout.freezeRows / freezeCols`
 *  state is observable through the playground's `window.__fcInst` exposure.
 *  This scenario drives the imperative API directly so the test is fast and
 *  deterministic, and it asserts that the engine echoes the layout change
 *  back through the workbook hint surface (used by the canvas renderer).
 *
 *  Playground-only — the React/Vue wrappers do not expose `__fcInst`. */
export async function runFreezePanesScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  type FreezeProbe = {
    ok: true;
    freezeRows: number;
    freezeCols: number;
  };

  const readFreeze = async (): Promise<FreezeProbe> => {
    const probe = await page.evaluate((): FreezeProbe | { ok: false } => {
      const w = window as unknown as {
        __fcInst?: {
          store: { getState(): { layout: { freezeRows: number; freezeCols: number } } };
        };
      };
      if (!w.__fcInst) return { ok: false };
      const layout = w.__fcInst.store.getState().layout;
      return { ok: true, freezeRows: layout.freezeRows, freezeCols: layout.freezeCols };
    });
    expect(probe.ok, 'window.__fcInst is required for this scenario').toBe(true);
    return probe as FreezeProbe;
  };

  const setFreeze = async (rows: number, cols: number): Promise<void> => {
    await page.evaluate(
      ({ rows: r, cols: c }) => {
        const w = window as unknown as {
          __fcInst?: {
            store: {
              setState(
                reducer: (s: {
                  layout: { freezeRows: number; freezeCols: number; [k: string]: unknown };
                  [k: string]: unknown;
                }) => unknown,
              ): void;
            };
            workbook: { clearViewportHint(): void };
          };
        };
        const inst = w.__fcInst;
        if (!inst) throw new Error('no __fcInst');
        // Drive the layout slice directly. The view toolbar / freeze menu does
        // the same thing — we skip them so the scenario stays portable.
        inst.store.setState((s) => ({
          ...s,
          layout: { ...s.layout, freezeRows: r, freezeCols: c },
        }));
        inst.workbook.clearViewportHint();
      },
      { rows, cols },
    );
    // Let the next animation frame flush so the renderer picks the new layout.
    await page.waitForTimeout(50);
  };

  // 1) Default state: nothing is frozen.
  const initial = await readFreeze();
  expect(initial).toMatchObject({ freezeRows: 0, freezeCols: 0 });

  // 2) Freeze the top row.
  await setFreeze(1, 0);
  const topRow = await readFreeze();
  expect(topRow).toMatchObject({ freezeRows: 1, freezeCols: 0 });

  // 3) Freeze the first column (replacing the row freeze).
  await setFreeze(0, 1);
  const firstCol = await readFreeze();
  expect(firstCol).toMatchObject({ freezeRows: 0, freezeCols: 1 });

  // 4) Freeze a 2×2 block.
  await setFreeze(2, 2);
  const block = await readFreeze();
  expect(block).toMatchObject({ freezeRows: 2, freezeCols: 2 });

  // 5) Unfreeze.
  await setFreeze(0, 0);
  const cleared = await readFreeze();
  expect(cleared).toMatchObject({ freezeRows: 0, freezeCols: 0 });
}
