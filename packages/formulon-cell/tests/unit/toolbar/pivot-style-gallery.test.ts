import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { createPivotTableAtPlacement } from '../../../src/commands/pivot-table.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import { createPivotStyleGallery } from '../../../src/toolbar/ribbon/pivot-style-gallery.js';

const setup = async () => {
  const wb = await WorkbookHandle.createDefault();
  for (const [r, row] of [
    ['Region', 'Sales'],
    ['East', 12],
    ['West', 5],
  ].entries()) {
    for (const [c, value] of row.entries()) {
      const addr = { sheet: 0, row: r, col: c };
      if (typeof value === 'number') wb.setNumber(addr, value);
      else wb.setText(addr, value);
    }
  }
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  const history = new History();
  createPivotTableAtPlacement(
    { origin: 'ribbon' as const, commandId: 'test', store, workbook: wb, history },
    { source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 }, rowField: 'Region', valueField: 'Sales' },
    { row: 8, col: 4 },
  );
  mutators.setActive(store, { sheet: 0, row: 9, col: 4 });
  const host = document.createElement('div');
  const inst = { host, store, history, workbook: wb } as unknown as SpreadsheetInstance;
  return { wb, store, history, inst };
};

describe('PivotTable style gallery', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('pages through None and the 84 built-in styles, six to a page', async () => {
    const { wb, inst } = await setup();
    const gallery = createPivotStyleGallery({ lang: 'ja', getInst: () => inst });
    document.body.appendChild(gallery);
    const labels = (): string[] =>
      Array.from(
        gallery.querySelectorAll('.fc-tb__pivotstyle-chip'),
        (c) => c.getAttribute('aria-label') ?? '',
      );
    // The gallery opens on the page holding the report's own style; go back to the first.
    const prev = gallery.querySelector<HTMLButtonElement>('.fc-tb__pivotstyle-prev');
    for (let i = 0; i < 20; i += 1) prev?.click();
    expect(labels()).toHaveLength(6);
    expect(labels()[0]).toBe('なし');
    expect(labels()[1]).toBe('白, ピボット スタイル (淡色) 1');
    const next = gallery.querySelector<HTMLButtonElement>('.fc-tb__pivotstyle-next');
    for (let i = 0; i < 2; i += 1) next?.click();
    // The third page holds None-relative items 13-18, which include Light 16.
    expect(labels()).toContain('アイスブルー, ピボット スタイル (淡色) 16');
    for (let i = 0; i < 20; i += 1) next?.click();
    expect(labels().at(-1)).toBe('オレンジ, ピボット スタイル (濃色) 28');
    expect(next?.disabled).toBe(true);
    wb.dispose();
  });

  it('marks the current style and applies a clicked one as a single undo step', async () => {
    const { wb, inst, history } = await setup();
    const gallery = createPivotStyleGallery({ lang: 'ja', getInst: () => inst });
    document.body.appendChild(gallery);
    const before = wb.getPivotTables()[0];
    if (!before) throw new Error('no pivot');
    for (let i = 0; i < 20; i += 1) {
      gallery.querySelector<HTMLButtonElement>('.fc-tb__pivotstyle-prev')?.click();
    }
    const chips = Array.from(
      gallery.querySelectorAll<HTMLButtonElement>('.fc-tb__pivotstyle-chip'),
    );
    chips[2]?.click();
    expect(wb.getPivotStyleInfo(before.sheetIndex, before.pivotIndex)?.name).toBe(
      'PivotStyleLight2',
    );
    const pressed = gallery.querySelector('.fc-tb__pivotstyle-chip[aria-pressed="true"]');
    expect(pressed?.getAttribute('aria-label')).toBe('白, ピボット スタイル (淡色) 2');
    expect(history.undo()).toBe(true);
    expect(wb.getPivotStyleInfo(before.sheetIndex, before.pivotIndex)?.name).not.toBe(
      'PivotStyleLight2',
    );
    wb.dispose();
  });

  it('draws each swatch with the style header band', async () => {
    const { wb, inst } = await setup();
    const gallery = createPivotStyleGallery({ lang: 'ja', getInst: () => inst });
    for (let i = 0; i < 20; i += 1) {
      gallery.querySelector<HTMLButtonElement>('.fc-tb__pivotstyle-prev')?.click();
    }
    const header = gallery
      .querySelectorAll('.fc-tb__pivotstyle-chip')[1]
      ?.querySelector<HTMLElement>('.fc-tb__pivotstyle-cell--header');
    expect(header?.style.background).not.toBe('');
    wb.dispose();
  });
});
