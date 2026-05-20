import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** C01 — internal copy/paste round-trip via Mod+C/V keyboard shortcuts.
 *  Canvas content isn't queryable, so we re-select each cell and read the
 *  formula bar to confirm the value landed. */
export async function runCopyPasteScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  // Seed A1=alpha; cursor advances to A2 on Enter.
  await sp.typeIntoActiveCell('alpha');
  // Step back to A1, copy.
  await page.keyboard.press('ArrowUp');
  await sp.shortcut('c');

  // Navigate to A3 and paste. Active cell after paste is the paste anchor.
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  await sp.shortcut('v');

  // Mod+V routes through navigator.clipboard.readText() which is async, so
  // poll the formula bar instead of asserting immediately.
  await expect.poll(() => sp.formulaBarValue(), { timeout: 2_000 }).toBe('alpha');
}

/** C02 — Mod+X cut → paste removes the source.
 *  Source becomes empty after paste; destination holds the value. */
export async function runCutPasteScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  await sp.typeIntoActiveCell('beta');
  await page.keyboard.press('ArrowUp');
  await sp.shortcut('x');

  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  await sp.shortcut('v');

  // Destination has the value (paste is async — see C01).
  await expect.poll(() => sp.formulaBarValue(), { timeout: 2_000 }).toBe('beta');

  // ...and the source (A1) is empty.
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('ArrowUp');
  expect(await sp.formulaBarValue()).toBe('');
}

/** C04 — multi-cell Mod+V paste must undo as one transaction. */
export async function runMultiCellPasteUndoScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  const readCell = (row: number, col: number) =>
    page.evaluate(
      ([r, c]: [number, number]) => {
        const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
          | { workbook: { getValue: (addr: { sheet: number; row: number; col: number }) => unknown } }
          | undefined;
        return inst?.workbook.getValue({ sheet: 0, row: r, col: c });
      },
      [row, col] as [number, number],
    );
  const seedCell = (row: number, col: number, value: string | number) =>
    page.evaluate(
      ([r, c, v]) => {
        const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
          | {
              workbook: {
                setNumber: (addr: { sheet: number; row: number; col: number }, value: number) => void;
                setText: (addr: { sheet: number; row: number; col: number }, value: string) => void;
                recalc: () => void;
              };
            }
          | undefined;
        if (!inst) return;
        const addr = { sheet: 0, row: r as number, col: c as number };
        if (typeof v === 'number') inst.workbook.setNumber(addr, v);
        else inst.workbook.setText(addr, v as string);
        inst.workbook.recalc();
      },
      [row, col, value],
    );

  await seedCell(0, 0, 'old-a');
  await seedCell(0, 1, 10);
  await seedCell(1, 0, 'old-b');
  await seedCell(1, 1, 20);
  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            setState: (updater: (state: { selection: unknown }) => unknown) => void;
          };
        }
      | undefined;
    inst?.store.setState((s) => ({
      ...s,
      selection: {
        ...(s.selection as object),
        active: { sheet: 0, row: 0, col: 0 },
        anchor: { sheet: 0, row: 0, col: 0 },
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      },
    }));
  });

  await expect.poll(() => readCell(0, 0), { timeout: 2_000 }).toEqual({
    kind: 'text',
    value: 'old-a',
  });
  await expect.poll(() => readCell(1, 1), { timeout: 2_000 }).toEqual({
    kind: 'number',
    value: 20,
  });

  await page.evaluate(() => navigator.clipboard.writeText('foo\t42\r\nbar\t99'));
  await sp.focusHost();
  await sp.shortcut('v');

  await expect.poll(() => readCell(0, 0), { timeout: 2_000 }).toEqual({
    kind: 'text',
    value: 'foo',
  });
  await expect.poll(() => readCell(1, 1), { timeout: 2_000 }).toEqual({
    kind: 'number',
    value: 99,
  });

  await sp.shortcut('z');

  await expect.poll(() => readCell(0, 0), { timeout: 2_000 }).toEqual({
    kind: 'text',
    value: 'old-a',
  });
  await expect.poll(() => readCell(0, 1), { timeout: 2_000 }).toEqual({
    kind: 'number',
    value: 10,
  });
  await expect.poll(() => readCell(1, 0), { timeout: 2_000 }).toEqual({
    kind: 'text',
    value: 'old-b',
  });
  await expect.poll(() => readCell(1, 1), { timeout: 2_000 }).toEqual({
    kind: 'number',
    value: 20,
  });

  await sp.shortcut('y');

  await expect.poll(() => readCell(0, 0), { timeout: 2_000 }).toEqual({
    kind: 'text',
    value: 'foo',
  });
  await expect.poll(() => readCell(0, 1), { timeout: 2_000 }).toEqual({
    kind: 'number',
    value: 42,
  });
  await expect.poll(() => readCell(1, 0), { timeout: 2_000 }).toEqual({
    kind: 'text',
    value: 'bar',
  });
  await expect.poll(() => readCell(1, 1), { timeout: 2_000 }).toEqual({
    kind: 'number',
    value: 99,
  });
}

/** C05 — ribbon Paste must use the same transaction semantics as Mod+V. */
export async function runRibbonPasteUndoScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  const readCell = (row: number, col: number) =>
    page.evaluate(
      ([r, c]: [number, number]) => {
        const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
          | { workbook: { getValue: (addr: { sheet: number; row: number; col: number }) => unknown } }
          | undefined;
        return inst?.workbook.getValue({ sheet: 0, row: r, col: c });
      },
      [row, col] as [number, number],
    );
  const seedCell = (row: number, col: number, value: string | number) =>
    page.evaluate(
      ([r, c, v]) => {
        const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
          | {
              workbook: {
                setNumber: (addr: { sheet: number; row: number; col: number }, value: number) => void;
                setText: (addr: { sheet: number; row: number; col: number }, value: string) => void;
                recalc: () => void;
              };
              store: {
                setState: (updater: (state: { selection: unknown }) => unknown) => void;
              };
            }
          | undefined;
        if (!inst) return;
        const addr = { sheet: 0, row: r as number, col: c as number };
        if (typeof v === 'number') inst.workbook.setNumber(addr, v);
        else inst.workbook.setText(addr, v as string);
        inst.workbook.recalc();
      },
      [row, col, value],
    );

  await seedCell(0, 0, 'old-a');
  await seedCell(0, 1, 10);
  await seedCell(1, 0, 'old-b');
  await seedCell(1, 1, 20);
  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            setState: (updater: (state: { selection: unknown }) => unknown) => void;
          };
        }
      | undefined;
    inst?.store.setState((s) => ({
      ...s,
      selection: {
        ...(s.selection as object),
        active: { sheet: 0, row: 0, col: 0 },
        anchor: { sheet: 0, row: 0, col: 0 },
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      },
    }));
  });

  await page.evaluate(() => navigator.clipboard.writeText('foo\t42\r\nbar\t99'));
  await page.locator('[data-ribbon-command="paste"]').click();

  await expect.poll(() => readCell(0, 0), { timeout: 2_000 }).toEqual({
    kind: 'text',
    value: 'foo',
  });
  await expect.poll(() => readCell(1, 1), { timeout: 2_000 }).toEqual({
    kind: 'number',
    value: 99,
  });

  await sp.shortcut('z');

  await expect.poll(() => readCell(0, 0), { timeout: 2_000 }).toEqual({
    kind: 'text',
    value: 'old-a',
  });
  await expect.poll(() => readCell(0, 1), { timeout: 2_000 }).toEqual({
    kind: 'number',
    value: 10,
  });
  await expect.poll(() => readCell(1, 0), { timeout: 2_000 }).toEqual({
    kind: 'text',
    value: 'old-b',
  });
  await expect.poll(() => readCell(1, 1), { timeout: 2_000 }).toEqual({
    kind: 'number',
    value: 20,
  });

  await sp.shortcut('y');

  await expect.poll(() => readCell(0, 0), { timeout: 2_000 }).toEqual({
    kind: 'text',
    value: 'foo',
  });
  await expect.poll(() => readCell(0, 1), { timeout: 2_000 }).toEqual({
    kind: 'number',
    value: 42,
  });
  await expect.poll(() => readCell(1, 0), { timeout: 2_000 }).toEqual({
    kind: 'text',
    value: 'bar',
  });
  await expect.poll(() => readCell(1, 1), { timeout: 2_000 }).toEqual({
    kind: 'number',
    value: 99,
  });
}
