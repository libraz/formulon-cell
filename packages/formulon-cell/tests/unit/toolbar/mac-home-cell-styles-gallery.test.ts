import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { addrKey } from '../../../src/engine/address.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import { createMacCellStylesGallery } from '../../../src/toolbar/ribbon/mac/home-styles-gallery.js';

const build = () => {
  const host = document.createElement('div');
  host.tabIndex = -1;
  document.body.appendChild(host);
  const store = createSpreadsheetStore();
  const history = new History();
  const inst = { host, store, history, workbook: null } as unknown as SpreadsheetInstance;
  const arrow = document.createElement('button');
  arrow.dataset.ribbonCommand = 'cellStyles';
  const menu = document.createElement('div');
  menu.className = 'fc-tb__menu';
  const gallery = createMacCellStylesGallery({
    lang: 'ja',
    getInst: () => inst,
    arrow,
    menu,
  });
  document.body.appendChild(gallery);
  return { gallery, store, arrow, menu };
};

describe('Mac Home cell styles gallery', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('shows the ten leading styles in two rows of five', () => {
    const { gallery } = build();
    const chips = Array.from(gallery.querySelectorAll<HTMLButtonElement>('[data-cell-style]'));
    expect(chips.map((c) => c.textContent)).toEqual([
      '標準',
      'どちらでもない',
      '悪い',
      '良い',
      'チェック セル',
      'メモ',
      'リンク セル',
      '計算',
      '警告文',
      '出力',
    ]);
    const grid = gallery.querySelector<HTMLElement>('.fc-tb__cellstyle-gallery-grid');
    expect(grid?.style.gridTemplateColumns).toBe('repeat(5, 1fr)');
    expect(chips[0]?.getAttribute('aria-pressed')).toBe('true');
  });

  it('puts the more arrow right after the swatches, followed by the existing menu', () => {
    const { gallery, arrow, menu } = build();
    const grid = gallery.querySelector('.fc-tb__cellstyle-gallery-grid');
    expect(grid?.nextElementSibling).toBe(arrow);
    expect(arrow.nextElementSibling).toBe(menu);
  });

  it('applies the clicked style to the selection', () => {
    const { gallery, store } = build();
    gallery.querySelector<HTMLButtonElement>('[data-cell-style="good"]')?.click();
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.cellStyle).toBeTruthy();
    expect(gallery.querySelector('[data-cell-style="good"]')?.getAttribute('aria-pressed')).toBe(
      'true',
    );
  });
});

describe('Mac Home cell styles gallery responsive states', () => {
  const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
  const css = readFileSync(join(root, 'src/styles/platform-mac.css'), 'utf8');
  const block = css.slice(css.indexOf('/* Mac Home cell styles gallery')).replace(/\s+/g, ' ');
  const rule = (selector: string): string => {
    const at = block.indexOf(selector);
    return at < 0 ? '' : block.slice(block.indexOf('{', at), block.indexOf('}', at));
  };

  it('keeps the strip and its fixed width in the wide ribbon', () => {
    expect(rule('.fc-tb__cellstyle-gallery {')).toContain('width: 430px');
    expect(rule('.fc-tb__cellstyle-gallery-grid {')).toContain('display: grid');
  });

  it('hides the strip when the group condenses and shows the plain button', () => {
    const condensed = '.fc-tb__ribbon-group--condensed';
    expect(rule(`${condensed} .fc-tb__cellstyle-gallery-grid`)).toContain('display: none');
    expect(rule(`${condensed} .fc-tb__cellstyle-gallery {`)).toContain('display: contents');
    expect(rule(`${condensed} .fc-tb__cellstyle-gallery-arrow`)).toContain('background: none');
    // The arrow's icon and label are hidden only outside a condensed group.
    expect(
      rule(
        '.fc-tb__ribbon-group:not(.fc-tb__ribbon-group--condensed) .fc-tb__cellstyle-gallery .fc-tb__cellstyle-gallery-arrow.fc-tb__rb > :is(span, svg)',
      ),
    ).toContain('display: none');
  });
});
