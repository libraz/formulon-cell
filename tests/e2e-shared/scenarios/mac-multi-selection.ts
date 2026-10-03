import { expect, type Page } from '@playwright/test';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

const selectedCells = (page: Page): Promise<string[]> =>
  page.evaluate(() => {
    const inst = (
      window as unknown as {
        __fcInst: {
          store: {
            getState(): {
              selection: {
                range: { r0: number; r1: number; c0: number; c1: number };
                extraRanges?: Array<{ r0: number; r1: number; c0: number; c1: number }>;
              };
            };
          };
        };
      }
    ).__fcInst;
    const selection = inst.store.getState().selection;
    const cells = new Set<string>();
    for (const range of [selection.range, ...(selection.extraRanges ?? [])]) {
      for (let row = range.r0; row <= range.r1; row++) {
        for (let col = range.c0; col <= range.c1; col++) cells.add(`${row}:${col}`);
      }
    }
    return [...cells].sort();
  });
const square = (size: number): string[] =>
  Array.from(
    { length: size * size },
    (_, index) => `${Math.floor(index / size)}:${index % size}`,
  ).sort();
const point = async (page: Page, row: number, col: number): Promise<{ x: number; y: number }> => {
  const box = await page.locator('.fc-host__grid').boundingBox();
  if (!box) throw new Error('Missing grid');
  return { x: box.x + 26 + 75 * (col + 0.5), y: box.y + 20 + 20 * (row + 0.5) };
};
const commandClick = async (page: Page, row: number, col: number): Promise<void> => {
  const cell = await point(page, row, col);
  await page.keyboard.down('Meta');
  await page.mouse.click(cell.x, cell.y);
  await page.keyboard.up('Meta');
};

export async function runMacSelectionToggle(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.enter('B2', 'keep');
  await sp.goTo('A1:C3');
  await commandClick(page, 1, 1);
  await expect.poll(() => selectedCells(page)).toEqual(square(3).filter((cell) => cell !== '1:1'));
  await expect(page.locator('.fc-host__statusbar-right')).toHaveText('8 cells');
  await expect.poll(() => sp.readValue('B2')).toEqual({ kind: 'text', value: 'keep' });
  await commandClick(page, 1, 1);
  await expect.poll(() => selectedCells(page)).toEqual(square(3));
  await commandClick(page, 3, 3);
  await expect.poll(() => selectedCells(page)).toEqual([...square(3), '3:3'].sort());
  await commandClick(page, 3, 3);
  await expect.poll(() => selectedCells(page)).toEqual(square(3));
  await sp.goTo('A1');
  await commandClick(page, 0, 0);
  await expect.poll(() => selectedCells(page)).toEqual(['0:0']);
  await expect(page.locator('.fc-host__statusbar-right')).toHaveText('1 cell');
  await sp.expectNoConsoleErrors();
}

export async function runMacSelectionDrag(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.goTo('A1:C3');
  const start = await point(page, 1, 1);
  const end = await point(page, 2, 2);
  await page.keyboard.down('Meta');
  await page.mouse.move(start.x, start.y);
  await page.mouse.down();
  await page.mouse.move(end.x, end.y, { steps: 6 });
  await expect.poll(() => selectedCells(page)).toEqual(['0:0', '0:1', '0:2', '1:0', '2:0']);
  // Shrinking the same gesture must restore cells removed by its earlier tip.
  await page.mouse.move(start.x, start.y, { steps: 6 });
  await expect.poll(() => selectedCells(page)).toEqual(square(3).filter((cell) => cell !== '1:1'));
  await page.mouse.up();
  await page.keyboard.up('Meta');
  await sp.goTo('A1:C3');
  const addStart = await point(page, 0, 3);
  const addEnd = await point(page, 1, 4);
  await page.keyboard.down('Meta');
  await page.mouse.move(addStart.x, addStart.y);
  await page.mouse.down();
  await page.mouse.move(addEnd.x, addEnd.y, { steps: 6 });
  await page.mouse.up();
  await page.keyboard.up('Meta');
  await expect
    .poll(() => selectedCells(page))
    .toEqual([...square(3), '0:3', '0:4', '1:3', '1:4'].sort());
  await sp.expectNoConsoleErrors();
}

export async function runMacMergedSelectionToggle(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.goTo('B2:C3');
  await sp.clickRibbon('merge');
  await sp.goTo('A1:D4');
  await commandClick(page, 2, 2);
  await expect
    .poll(() => selectedCells(page))
    .toEqual(square(4).filter((cell) => !['1:1', '1:2', '2:1', '2:2'].includes(cell)));
  await commandClick(page, 2, 2);
  await expect.poll(() => selectedCells(page)).toEqual(square(4));
  await sp.expectNoConsoleErrors();
}

export async function runMacSelectionInputWithHole(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.enter('B2', 'keep');
  for (const input of ['cell', 'formulaBar']) {
    await sp.goTo('A1:C3');
    await commandClick(page, 1, 1);
    await expect
      .poll(() => selectedCells(page))
      .toEqual(square(3).filter((cell) => cell !== '1:1'));
    if (input === 'cell') {
      await page.keyboard.type('7');
      await page.keyboard.press('Meta+Enter');
    } else {
      const fx = page.locator('.fc-host__formulabar-input');
      await fx.fill('7');
      await fx.press('Meta+Enter');
    }
    for (const ref of ['A1', 'B1', 'C1', 'A2', 'C2', 'A3', 'B3', 'C3'])
      await expect.poll(() => sp.readValue(ref)).toEqual({ kind: 'number', value: 7 });
    await expect.poll(() => sp.readValue('B2')).toEqual({ kind: 'text', value: 'keep' });
    await page.locator('.fc-host').focus();
    await page.keyboard.press('Meta+z');
    for (const ref of ['A1', 'B1', 'C1', 'A2', 'C2', 'A3', 'B3', 'C3'])
      await expect.poll(() => sp.readValue(ref)).toEqual({ kind: 'blank' });
    await expect.poll(() => sp.readValue('B2')).toEqual({ kind: 'text', value: 'keep' });
  }
  await sp.expectNoConsoleErrors();
}

export async function runMacHeaderSelectionToggle(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  const headerPoint = async (axis: 'row' | 'col', index: number) => {
    const box = await page.locator('.fc-host__grid').boundingBox();
    if (!box) throw new Error('Missing grid');
    return axis === 'row'
      ? { x: box.x + 12, y: box.y + 20 + 20 * (index + 0.5) }
      : { x: box.x + 26 + 75 * (index + 0.5), y: box.y + 10 };
  };
  const bands = (axis: 'row' | 'col') =>
    page.evaluate((direction) => {
      const inst = (
        window as unknown as {
          __fcInst: {
            store: {
              getState(): {
                selection: {
                  range: { r0: number; r1: number; c0: number; c1: number };
                  extraRanges?: Array<{ r0: number; r1: number; c0: number; c1: number }>;
                };
              };
            };
          };
        }
      ).__fcInst;
      const s = inst.store.getState().selection;
      return [s.range, ...(s.extraRanges ?? [])]
        .map((r) => (direction === 'row' ? [r.r0, r.r1, r.c0, r.c1] : [r.c0, r.c1, r.r0, r.r1]))
        .sort((a, b) => (a[0] ?? 0) - (b[0] ?? 0));
    }, axis);
  for (const axis of ['row', 'col'] as const) {
    const last = axis === 'row' ? 16_383 : 1_048_575;
    const first = await headerPoint(axis, 0);
    const far = await headerPoint(axis, 3);
    await page.mouse.click(first.x, first.y);
    await page.keyboard.down('Shift');
    await page.mouse.click(far.x, far.y);
    await page.keyboard.up('Shift');
    await expect.poll(() => bands(axis)).toEqual([[0, 3, 0, last]]);
    const start = await headerPoint(axis, 1);
    const tip = await headerPoint(axis, 2);
    await page.keyboard.down('Meta');
    await page.mouse.move(start.x, start.y);
    await page.mouse.down();
    await page.mouse.move(tip.x, tip.y, { steps: 6 });
    await expect
      .poll(() => bands(axis))
      .toEqual([
        [0, 0, 0, last],
        [3, 3, 0, last],
      ]);
    await page.mouse.move(start.x, start.y, { steps: 6 });
    await expect
      .poll(() => bands(axis))
      .toEqual([
        [0, 0, 0, last],
        [2, 3, 0, last],
      ]);
    await page.mouse.up();
    await page.mouse.click(start.x, start.y);
    await page.keyboard.up('Meta');
    await expect
      .poll(() => bands(axis))
      .toEqual([
        [0, 0, 0, last],
        [1, 1, 0, last],
        [2, 3, 0, last],
      ]);
  }
  await sp.goTo('A1:B2');
  const firstRow = await headerPoint('row', 0);
  await page.keyboard.down('Meta');
  await page.mouse.click(firstRow.x, firstRow.y);
  await page.keyboard.up('Meta');
  await expect
    .poll(() => bands('row'))
    .toEqual([
      [0, 0, 0, 16_383],
      [1, 1, 0, 1],
    ]);
  await sp.expectNoConsoleErrors();
}

export async function runMacSelectionClearWithHole(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await sp.goTo('A1:C3');
  await page.keyboard.type('7');
  await page.keyboard.press('Meta+Enter');
  await sp.enter('B2', 'keep');
  await sp.enter('C3', '=1+2');
  await sp.enter('D4', '99');
  const targets = ['A1', 'B1', 'C1', 'A2', 'C2', 'A3', 'B3', 'C3'];
  for (const action of ['Delete', 'ribbon']) {
    await sp.goTo('A1:C3');
    await commandClick(page, 1, 1);
    await expect
      .poll(() => selectedCells(page))
      .toEqual(square(3).filter((cell) => cell !== '1:1'));
    if (action === 'Delete') await page.keyboard.press('Delete');
    else {
      await sp.clickRibbon('clearFormat');
      const clear = page.locator('#menu-clear [data-clear="contents"]');
      await expect(clear).toBeEnabled();
      await clear.click();
    }
    for (const ref of targets)
      await expect.poll(() => sp.readValue(ref)).toEqual({ kind: 'blank' });
    await expect.poll(() => sp.readValue('B2')).toEqual({ kind: 'text', value: 'keep' });
    await expect.poll(() => sp.readValue('D4')).toEqual({ kind: 'number', value: 99 });
    await page.locator('.fc-host').focus();
    await page.keyboard.press('Meta+z');
    for (const ref of targets)
      await expect
        .poll(() => sp.readValue(ref))
        .toEqual({ kind: 'number', value: ref === 'C3' ? 3 : 7 });
    await expect.poll(() => sp.readFormula('C3')).toBe('=1+2');
    await page.keyboard.press('Meta+Shift+z');
    for (const ref of targets)
      await expect.poll(() => sp.readValue(ref)).toEqual({ kind: 'blank' });
    await page.keyboard.press('Meta+z');
  }
  // The active primary can be blank while an additional selected area has content.
  await sp.goTo('A1');
  await commandClick(page, 4, 4);
  await sp.clickRibbon('clearFormat');
  const clear = page.locator('#menu-clear [data-clear="contents"]');
  await expect(clear).toBeEnabled();
  await clear.click();
  await expect.poll(() => sp.readValue('A1')).toEqual({ kind: 'blank' });
  await expect.poll(() => sp.readValue('B2')).toEqual({ kind: 'text', value: 'keep' });
  await sp.expectNoConsoleErrors();
}
