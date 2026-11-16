import { afterEach, describe, expect, it, vi } from 'vitest';
import { canExecuteBuiltIn } from '../../../src/commands/built-in-command-policy.js';
import { customPivotTableStyleId } from '../../../src/commands/format-as-table.js';
import { History } from '../../../src/commands/history.js';
import { createPivotTableAtPlacement } from '../../../src/commands/pivot-table.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { dictionaries } from '../../../src/i18n/strings.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import {
  createPivotStyleGallery,
  pivotStyleLabel,
} from '../../../src/toolbar/ribbon/pivot-style-gallery.js';

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

// The accessible names the Japanese app gives the built-in styles: "<colour>, <name>".
const JA_FAMILY = { Light: '淡色', Medium: '中間', Dark: '濃色' } as const;
const JA_COLOURS: Record<keyof typeof JA_FAMILY, readonly string[]> = {
  Light: [
    ...Array<string>(8).fill('白'),
    ...[
      'アイスブルー',
      'ローズ',
      '薄い緑',
      'ラベンダー',
      'スカイブルー',
      '薄いオレンジ',
      '薄い灰色',
    ],
    ...['アイスブルー', 'ローズ', '薄い緑', 'ラベンダー', 'スカイブルー', '薄いオレンジ'],
    ...Array<string>(7).fill('白'),
  ],
  Medium: [
    ...['白', '薄い青', 'ローズ', '薄い緑', 'ラベンダー', 'スカイブルー', '薄いオレンジ'],
    ...['白', 'アイスブルー', 'ローズ', '薄い緑', 'ラベンダー', 'スカイブルー', '薄いオレンジ'],
    ...['白', 'アイスブルー', 'ローズ', '薄い緑', 'ラベンダー', 'スカイブルー', '薄いオレンジ'],
    ...[
      '薄い灰色',
      'アイスブルー',
      'ローズ',
      '薄い緑',
      'ラベンダー',
      'スカイブルー',
      '薄いオレンジ',
    ],
  ],
  Dark: [
    ...['薄い灰色', '濃い青', '濃い赤', '濃い緑', '濃い紫', '濃い青緑', '茶'],
    ...Array<string>(7).fill('濃い灰色'),
    ...['灰色', '青', '赤', 'オリーブ', '濃い紫', 'アクア', 'オレンジ'],
    ...['灰色', '青', '赤', 'オリーブ', '濃い紫', 'アクア', 'オレンジ'],
  ],
};

describe('PivotTable style names', () => {
  it('keeps the Japanese accessible name of every built-in style', () => {
    expect(pivotStyleLabel('', 'ja')).toBe('なし');
    for (const family of ['Light', 'Medium', 'Dark'] as const) {
      for (let n = 1; n <= 28; n += 1) {
        expect(pivotStyleLabel(`PivotStyle${family}${n}`, 'ja')).toBe(
          `${JA_COLOURS[family][n - 1]}, ピボット スタイル (${JA_FAMILY[family]}) ${n}`,
        );
      }
    }
  });

  it('names the styles in German and English with the app strings and the same pattern', () => {
    expect(pivotStyleLabel('', 'de')).toBe('Keine');
    expect(pivotStyleLabel('PivotStyleLight1', 'de')).toBe('Weiß, Pivotformat - Hell 1');
    expect(pivotStyleLabel('PivotStyleDark28', 'de')).toBe('Orange, Pivotformat - Dunkel 28');
    expect(pivotStyleLabel('PivotStyleLight1', 'en')).toBe('White, Pivot Style Light 1');
  });

  it('leaves a name outside the built-in set as it is', () => {
    expect(pivotStyleLabel('PivotStyleLight29', 'de')).toBe('PivotStyleLight29');
    expect(pivotStyleLabel('Brand', 'de')).toBe('Brand');
  });
});

describe('PivotTable style gallery', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('pages through None and the 84 built-in styles, six to a page', async () => {
    const { wb, inst } = await setup();
    const gallery = createPivotStyleGallery({
      lang: 'ja',
      getInst: () => inst,
      arrow: document.createElement('button'),
      menu: null,
      mac: false,
    });
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

  // Excel for Mac: seven 73 x 56 chips stepping 71, no page arrows, scrolled so the report's
  // style is the fourth chip.
  it('scrolls the Mac strip to keep the current style fourth of seven', async () => {
    const { wb, inst } = await setup();
    const gallery = createPivotStyleGallery({
      lang: 'ja',
      getInst: () => inst,
      arrow: document.createElement('button'),
      menu: null,
      mac: true,
    });
    document.body.appendChild(gallery);
    const labels = Array.from(
      gallery.querySelectorAll('.fc-tb__pivotstyle-chip'),
      (c) => c.getAttribute('aria-label') ?? '',
    );
    expect(labels).toHaveLength(7);
    expect(labels[0]).toContain('(淡色) 13');
    expect(labels[3]).toContain('(淡色) 16');
    expect(labels[6]).toContain('(淡色) 19');
    expect(gallery.querySelector('.fc-tb__pivotstyle-prev')).toBeNull();
    expect(
      ['--fc-strip-chip-w', '--fc-strip-chip-h', '--fc-strip-step', '--fc-strip-visible'].map(
        (name) => gallery.style.getPropertyValue(name),
      ),
    ).toEqual(['73px', '56px', '71px', '7']);
    wb.dispose();
  });

  // The chip count follows the strip's width: as many 71-step chips as fit, the current style
  // fourth while more than four fit and last otherwise. The strip re-renders a frame after the
  // resize, never inside the observer callback (a resize loop in WebKit).
  it('shows the chips the Mac strip width fits, keeping the current style in place', async () => {
    const observers: (() => void)[] = [];
    const frames: FrameRequestCallback[] = [];
    vi.stubGlobal('requestAnimationFrame', (cb: FrameRequestCallback) => frames.push(cb));
    vi.stubGlobal(
      'ResizeObserver',
      class {
        constructor(cb: () => void) {
          observers.push(cb);
        }
        observe(): void {}
        disconnect(): void {}
      },
    );
    try {
      const { wb, inst } = await setup();
      const gallery = createPivotStyleGallery({
        lang: 'ja',
        getInst: () => inst,
        arrow: document.createElement('button'),
        menu: null,
        mac: true,
      });
      document.body.appendChild(gallery);
      const strip = gallery.querySelector<HTMLElement>('.fc-tb__pivotstyle-strip');
      if (!strip) throw new Error('no strip');
      const labels = (): string[] =>
        Array.from(
          gallery.querySelectorAll('.fc-tb__pivotstyle-chip'),
          (c) => c.getAttribute('aria-label') ?? '',
        );
      const resize = (width: number): string[] => {
        Object.defineProperty(strip, 'clientWidth', { configurable: true, value: width });
        const before = labels();
        for (const cb of observers) cb();
        expect(labels()).toEqual(before);
        for (const frame of frames.splice(0)) frame(0);
        return labels();
      };
      const five = resize(5 * 71 + 10);
      expect(five).toHaveLength(5);
      expect(five[3]).toContain('(淡色) 16');
      const two = resize(2 * 71);
      expect(two).toHaveLength(2);
      expect(two[1]).toContain('(淡色) 16');
      expect(resize(40)).toHaveLength(1);
      expect(resize(2000)).toHaveLength(7);
      wb.dispose();
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('marks the current style and applies a clicked one as a single undo step', async () => {
    const { wb, inst, history } = await setup();
    const gallery = createPivotStyleGallery({
      lang: 'ja',
      getInst: () => inst,
      arrow: document.createElement('button'),
      menu: null,
      mac: false,
    });
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

  // The strip follows the restyle's permission, as the table and sparkline strips do.
  it('greys the strip while the restyle is refused', async () => {
    const { wb, store, inst } = await setup();
    Object.assign(inst, { i18n: { strings: dictionaries.en } });
    const gallery = createPivotStyleGallery({
      lang: 'en',
      getInst: () => inst,
      arrow: document.createElement('button'),
      menu: null,
      mac: false,
    });
    document.body.appendChild(gallery);
    const chips = () =>
      Array.from(gallery.querySelectorAll<HTMLButtonElement>('.fc-tb__pivotstyle-chip'));
    const refresh = () => gallery.dispatchEvent(new Event('fc-ribbon-control-refresh'));
    expect(chips().every((chip) => !chip.disabled)).toBe(true);
    setProtectedSheet(store, 0, true);
    refresh();
    expect(chips().length).toBeGreaterThan(0);
    expect(chips().every((chip) => chip.disabled)).toBe(true);
    // With PivotTables allowed, the strip follows the restyle's own permission.
    setProtectedSheet(store, 0, true, { permissions: { pivotTables: true } });
    refresh();
    const allowed = canExecuteBuiltIn(store, 'pivotQuickStyles', 'ribbon').allowed;
    expect(chips().every((chip) => chip.disabled === !allowed)).toBe(true);
    setProtectedSheet(store, 0, false);
    refresh();
    expect(chips().every((chip) => !chip.disabled)).toBe(true);
    wb.dispose();
  });

  it('draws each swatch with the style header band', async () => {
    const { wb, inst } = await setup();
    const gallery = createPivotStyleGallery({
      lang: 'ja',
      getInst: () => inst,
      arrow: document.createElement('button'),
      menu: null,
      mac: false,
    });
    for (let i = 0; i < 20; i += 1) {
      gallery.querySelector<HTMLButtonElement>('.fc-tb__pivotstyle-prev')?.click();
    }
    const header = gallery
      .querySelectorAll('.fc-tb__pivotstyle-chip')[1]
      ?.querySelector<HTMLElement>('.fc-tb__pivotstyle-cell--header');
    expect(header?.style.background).not.toBe('');
    wb.dispose();
  });

  it('lists custom PivotTable styles first and applies one', async () => {
    const { wb, inst, store } = await setup();
    mutators.upsertCustomPivotTableStyle(store, {
      id: customPivotTableStyleId('Brand'),
      label: 'Brand',
      elements: { headerRow: { fill: '#336699' } },
    });
    const gallery = createPivotStyleGallery({
      lang: 'ja',
      getInst: () => inst,
      arrow: document.createElement('button'),
      menu: null,
      mac: false,
    });
    for (let i = 0; i < 20; i += 1) {
      gallery.querySelector<HTMLButtonElement>('.fc-tb__pivotstyle-prev')?.click();
    }
    const chips = Array.from(
      gallery.querySelectorAll<HTMLButtonElement>('.fc-tb__pivotstyle-chip'),
    );
    expect(chips.map((c) => c.getAttribute('aria-label')).slice(0, 2)).toEqual(['Brand', 'なし']);
    chips[0]?.click();
    expect(store.getState().tables.pivotTableStyles).toEqual([
      { sheetIndex: 0, pivotIndex: 0, styleId: customPivotTableStyleId('Brand') },
    ]);
    // A built-in pick replaces the custom assignment.
    gallery.querySelectorAll<HTMLButtonElement>('.fc-tb__pivotstyle-chip')[2]?.click();
    expect(store.getState().tables.pivotTableStyles).toEqual([]);
    expect(wb.getPivotStyleInfo(0, 0)?.name).toBe('PivotStyleLight1');
    wb.dispose();
  });
});
