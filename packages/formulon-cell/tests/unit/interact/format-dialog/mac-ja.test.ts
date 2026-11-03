import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { ja } from '../../../../src/i18n/strings/ja.js';
import { attachFormatDialog } from '../../../../src/interact/format-dialog.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';
import { setActive, setRange } from './fixtures.js';

const mount = (platform: 'mac' | 'default') => {
  const fcHost = document.createElement('div');
  fcHost.className = 'fc-host';
  fcHost.dataset.fcPlatform = platform;
  const host = document.createElement('div');
  host.tabIndex = -1;
  fcHost.appendChild(host);
  document.body.appendChild(fcHost);
  const store = createSpreadsheetStore();
  setActive(store, 0, 0);
  const handle = attachFormatDialog({ host, store, strings: ja });
  return handle;
};

const text = (selector: string): string[] =>
  [...document.querySelectorAll(selector)].map((el) => el.textContent ?? '');

afterEach(() => {
  document.body.innerHTML = '';
});

describe('format dialog, Mac Japanese presentation', () => {
  let handle: ReturnType<typeof mount>;
  beforeEach(() => {
    handle = mount('mac');
  });

  it('uses the spreadsheet app wording for the category descriptions', () => {
    handle.open('number');
    expect(document.querySelector('.fc-fmtdlg__hintbar')?.textContent).toBe(
      'セルの値に対して一般の書式を適用します (特定の書式を指定しません)。',
    );
    document
      .querySelector<HTMLButtonElement>('button[data-fc-cat="fixed"]')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(document.querySelector('.fc-fmtdlg__hintbar')?.textContent).toBe(
      '数値の表示形式を設定します。小数点位置を揃える場合は、[会計] を選択してください。',
    );
    expect(text('.fc-fmtdlg__row span, .fc-fmtdlg__check span')).toEqual(
      expect.arrayContaining(['小数点以下の桁数:', '桁区切り ( , ) を使用する']),
    );
  });

  it('labels the alignment, border and fill controls as the app does', () => {
    handle.open('align');
    const labels = text('.fc-fmtdlg__panel span, .fc-fmtdlg__panel div');
    for (const label of [
      '水平方向の配置:',
      '垂直方向の配置:',
      'インデント:',
      '文字の制御:',
      '方向:',
      '度:',
    ]) {
      expect(labels).toContain(label);
    }
    handle.close();
    handle.open('border');
    expect(text('.fc-fmtdlg__panel span')).toEqual(
      expect.arrayContaining(['線のスタイル:', '線の色:']),
    );
    handle.close();
    handle.open('fill');
    expect(text('.fc-fmtdlg__panel span')).toEqual(
      expect.arrayContaining(['背景色:', 'パターンの色:', 'パターン スタイル:']),
    );
  });
});

describe('format dialog alignment tab on the Mac profile', () => {
  it('drops the text-direction row only there', () => {
    mount('mac').open('align');
    const row = document.querySelector<HTMLElement>('.fc-fmtdlg__text-direction-row');
    expect(row?.hidden).toBe(true);
    document.body.innerHTML = '';
    mount('default').open('align');
    expect(document.querySelector<HTMLElement>('.fc-fmtdlg__text-direction-row')?.hidden).toBe(
      false,
    );
  });
});

describe('format dialog pattern lists', () => {
  const itemsFor = (platform: 'mac' | 'default', cat: string): number => {
    const handle = mount(platform);
    handle.open('number');
    document
      .querySelector<HTMLButtonElement>(`button[data-fc-cat="${cat}"]`)
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    const count = document.querySelectorAll(
      '.fc-fmtdlg__cat-controls .fc-fmtdlg__pattern-list button',
    ).length;
    document.body.innerHTML = '';
    return count;
  };

  it('lists the 19 date and 9 time formats of the Mac app, 5 and 4 elsewhere', () => {
    expect(itemsFor('mac', 'date')).toBe(19);
    expect(itemsFor('mac', 'time')).toBe(9);
    expect(itemsFor('default', 'date')).toBe(5);
    expect(itemsFor('default', 'time')).toBe(4);
  });
});

describe('format dialog negative-number lists', () => {
  const rows = (platform: 'mac' | 'default', cat: string): string[] => {
    const handle = mount(platform);
    handle.open('number');
    document
      .querySelector<HTMLButtonElement>(`button[data-fc-cat="${cat}"]`)
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    const out = [...document.querySelectorAll<HTMLElement>('[data-fc-negative-style]')]
      .filter((el) => !el.hidden)
      .map((el) => `${el.dataset.fcNegativeStyle}:${el.textContent}`);
    document.body.innerHTML = '';
    return out;
  };

  it('lists 7 number and 5 currency rows in the Mac app order', () => {
    expect(rows('mac', 'fixed')).toEqual([
      'red-parens:(1234)',
      'parens:(1234)',
      'red-plain:1234',
      'minus:-1234',
      'red:-1234',
      'triangle:△ 1234',
      'triangle-filled:▲ 1234',
    ]);
    expect(rows('mac', 'currency').map((row) => row.split(':')[0])).toEqual([
      'red-parens',
      'parens',
      'red-plain',
      'minus',
      'red',
    ]);
  });

  it('keeps the four shared rows elsewhere', () => {
    expect(rows('default', 'fixed').map((row) => row.split(':')[0])).toEqual([
      'parens',
      'red-parens',
      'red',
      'minus',
    ]);
  });
});

describe('format dialog fill tab on the Mac profile', () => {
  const click = (el: Element | null | undefined): void => {
    el?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
  };
  const popupLabels = (): string[] =>
    [...document.querySelectorAll('.fc-fmtdlg__fill-popup-label')].map(
      (el) => el.textContent ?? '',
    );

  it('shows colour and pattern pickers as popups with No colour / Automatic defaults', () => {
    mount('mac').open('fill');
    expect(popupLabels()).toEqual(['色なし', '自動', '']);
    expect(document.querySelector<HTMLElement>('[data-fc-fill-effects]')?.hidden).toBe(true);
  });

  it('opens the background popup, picks a swatch and returns to No colour', () => {
    mount('mac').open('fill');
    const popup = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__fill-popup');
    const flyout = document.querySelector<HTMLElement>('.fc-fmtdlg__fill-popup-flyout');
    expect(flyout?.hidden).toBe(true);
    click(popup);
    expect(flyout?.hidden).toBe(false);
    click(flyout?.querySelector('[data-color]'));
    expect(flyout?.hidden).toBe(true);
    expect(popupLabels()[0]).toMatch(/^#[0-9a-f]{6}$/);
    click(popup);
    click(flyout?.querySelector('[data-fc-fill-none]'));
    expect(popupLabels()[0]).toBe('色なし');
  });

  it('keeps the shared fill controls on other platforms', () => {
    mount('default').open('fill');
    expect(document.querySelector('.fc-fmtdlg__fill-popup')).toBeNull();
  });
});

describe('format dialog border tab on the Mac profile', () => {
  const click = (el: Element | null | undefined): void => {
    el?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
  };

  it('adds the preset and outline headings, the hint and the No line cell', () => {
    mount('mac').open('border');
    const text = text2('.fc-fmtdlg__border-heading');
    expect(text).toEqual(expect.arrayContaining(['プリセット:', '輪郭:']));
    expect(document.querySelector('.fc-fmtdlg__border-hint')?.textContent).toContain(
      'プレビュー ボックス内をクリックして罫線を適用することもできます。',
    );
    const cells = [...document.querySelectorAll('.fc-fmtdlg__line-style')];
    expect(cells).toHaveLength(14);
    expect(cells[0]?.getAttribute('data-border-style')).toBe('none');
  });

  it('lays the outline out as the Mac app does: 文字列 box, inside buttons disabled on one cell', () => {
    mount('mac').open('border');
    expect(document.querySelector('.fc-fmtdlg__border-preview')?.textContent).toBe('文字列');
    expect(document.querySelectorAll('.fc-fmtdlg__border-tick')).toHaveLength(4);
    const inside = [...document.querySelectorAll<HTMLButtonElement>('[data-border-inside]')];
    expect(inside.map((b) => b.dataset.borderInside)).toEqual(['h', 'v']);
    expect(inside.every((b) => b.disabled)).toBe(true);
    expect(
      document.querySelector<HTMLButtonElement>('.fc-fmtdlg__border-preset--inside')?.disabled,
    ).toBe(true);
  });

  it('keeps the shared outline preview elsewhere', () => {
    mount('default').open('border');
    expect(document.querySelectorAll('.fc-fmtdlg__border-tick')).toHaveLength(0);
    expect(document.querySelectorAll('[data-border-inside]')).toHaveLength(0);
  });

  it('shows the line colour as a popup that starts on Automatic', () => {
    mount('mac').open('border');
    expect(
      document.querySelector('.fc-fmtdlg__border-color-row .fc-fmtdlg__color-toggle')?.textContent,
    ).toContain('自動');
  });

  it('clears an edge instead of drawing one while the No line cell is active', () => {
    mount('mac').open('border');
    const top = () =>
      document.querySelector<HTMLInputElement>('[data-fc-check="borderTop"]')?.checked;
    click(document.querySelector('[data-border-side="top"]'));
    expect(top()).toBe(true);
    click(document.querySelector('[data-border-style="none"]'));
    click(document.querySelector('[data-border-side="top"]'));
    expect(top()).toBe(false);
    click(document.querySelector('[data-border-side="bottom"]'));
    expect(
      document.querySelector<HTMLInputElement>('[data-fc-check="borderBottom"]')?.checked,
    ).toBe(false);
  });
});

const text2 = (selector: string): string[] =>
  [...document.querySelectorAll(selector)].map((el) => el.textContent ?? '');

describe('format dialog inside borders on a 2x2 selection (Mac)', () => {
  const run = (axes: Array<'h' | 'v'>) => {
    const fcHost = document.createElement('div');
    fcHost.className = 'fc-host';
    fcHost.dataset.fcPlatform = 'mac';
    const host = document.createElement('div');
    host.tabIndex = -1;
    fcHost.appendChild(host);
    document.body.appendChild(fcHost);
    const store = createSpreadsheetStore();
    setRange(store, 0, 0, 1, 1);
    const handle = attachFormatDialog({ host, store, strings: ja });
    handle.open('border');
    for (const axis of axes) {
      document
        .querySelector(`[data-border-inside="${axis}"]`)
        ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    }
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    const sides = (row: number, col: number) => {
      const key = `0:${row}:${col}`;
      return Object.keys(store.getState().format.formats.get(key)?.borders ?? {}).sort();
    };
    return { top: [sides(0, 0), sides(0, 1)], bottom: [sides(1, 0), sides(1, 1)] };
  };

  it('inside-horizontal alone draws only the middle horizontal edge', async () => {
    const out = run(['h']);
    await Promise.resolve();
    expect(out.top).toEqual([[], []]);
    expect(out.bottom).toEqual([['top'], ['top']]);
  });

  it('inside-vertical alone draws only the middle vertical edge', async () => {
    const out = run(['v']);
    await Promise.resolve();
    expect(out.top).toEqual([[], ['left']]);
    expect(out.bottom).toEqual([[], ['left']]);
  });

  it('both axes together are the Inside preset', async () => {
    const out = run(['h', 'v']);
    await Promise.resolve();
    expect(out.bottom).toEqual([['top'], ['left', 'top']]);
  });
});

describe('format dialog font tab on the Mac profile', () => {
  const click = (el: Element | null | undefined): void => {
    el?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
  };
  const visibleText = (selector: string): string[] =>
    [...document.querySelectorAll<HTMLElement>(selector)]
      .filter((el) => !el.closest('[hidden]'))
      .map((el) => el.textContent ?? '');

  it('labels the columns, shows a style field and a 文字飾り group of three', () => {
    mount('mac').open('font');
    const labels = visibleText(
      '.fc-fmtdlg__panel-tab[data-fc-tab="font"] span, .fc-fmtdlg__panel-tab[data-fc-tab="font"] div, .fc-fmtdlg__panel-tab[data-fc-tab="font"] label',
    );
    for (const label of ['フォント名:', 'スタイル:', 'サイズ:', '文字飾り:']) {
      expect(labels.some((text) => text.startsWith(label))).toBe(true);
    }
    const style = document.querySelector<HTMLInputElement>('.fc-fmtdlg__font-style-field');
    expect(style?.readOnly).toBe(true);
    expect(style?.value).toBe('標準');
    const checks = visibleText('.fc-fmtdlg__font-deco .fc-fmtdlg__check');
    expect(checks).toEqual(['取り消し線', '上付き', '下付き']);
  });

  it('drops the bold, italic, outline and shadow checkboxes (the style list covers them)', () => {
    mount('mac').open('font');
    for (const key of ['bold', 'italic', 'fontOutline', 'fontShadow']) {
      const box = document.querySelector<HTMLInputElement>(`input[data-fc-check="${key}"]`);
      expect(box?.closest('[hidden]')).not.toBeNull();
    }
  });

  it('puts Automatic inside the colour popup and keeps the style field in step', () => {
    mount('mac').open('font');
    const flyout = document.querySelector<HTMLElement>(
      '.fc-fmtdlg__font-color-row .fc-fmtdlg__color-flyout, .fc-fmtdlg__color-flyout',
    );
    expect(
      [
        ...document.querySelectorAll(
          '.fc-fmtdlg__panel-tab[data-fc-tab="font"] .fc-fmtdlg__color-flyout .fc-fmtdlg__btn',
        ),
      ].map((b) => b.textContent),
    ).toContain('自動');
    click(document.querySelector('[data-fc-font-style="boldItalic"]'));
    expect(document.querySelector<HTMLInputElement>('.fc-fmtdlg__font-style-field')?.value).toBe(
      '太字 斜体',
    );
    expect(flyout).not.toBeNull();
  });

  it('keeps the shared font controls elsewhere', () => {
    mount('default').open('font');
    expect(document.querySelector('.fc-fmtdlg__font-style-field')).toBeNull();
    expect(document.querySelector('.fc-fmtdlg__font-deco')).toBeNull();
  });
});

describe('format dialog open state on the Mac profile', () => {
  const click = (el: Element | null | undefined): void => {
    el?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
  };
  const selectedTab = (): string | null =>
    document.querySelector('.fc-fmtdlg__tab[aria-selected="true"]')?.getAttribute('data-fc-tab') ??
    null;

  it('has the six tabs of the Mac app (no その他) and reopens on the last used one', () => {
    const handle = mount('mac');
    handle.open();
    const tabs = [...document.querySelectorAll<HTMLElement>('.fc-fmtdlg__tab')]
      .filter((el) => !el.hidden)
      .map((el) => el.textContent);
    expect(tabs).toEqual(['表示形式', '配置', 'フォント', '罫線', '塗りつぶし', '保護']);
    expect(selectedTab()).toBe('number');
    click(document.querySelector('[data-fc-tab="protection"]'));
    handle.close();
    handle.open();
    expect(selectedTab()).toBe('protection');
  });

  it('always starts on the first tab elsewhere', () => {
    const handle = mount('default');
    handle.open();
    click(document.querySelector('[data-fc-tab="protection"]'));
    handle.close();
    handle.open();
    expect(selectedTab()).toBe('number');
    expect(
      [...document.querySelectorAll<HTMLElement>('.fc-fmtdlg__tab')].filter((el) => !el.hidden),
    ).toHaveLength(7);
  });
});

describe('format dialog number defaults', () => {
  const decimalsAfterPicking = (platform: 'mac' | 'default'): string | undefined => {
    const handle = mount(platform);
    handle.open('number');
    document
      .querySelector<HTMLButtonElement>('button[data-fc-cat="fixed"]')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    return document.querySelector<HTMLInputElement>('.fc-fmtdlg__cat-controls input[type="number"]')
      ?.value;
  };

  it('starts a category at 0 decimal places on the Mac profile, 2 elsewhere', () => {
    expect(decimalsAfterPicking('mac')).toBe('0');
    document.body.innerHTML = '';
    expect(decimalsAfterPicking('default')).toBe('2');
  });
});

describe('format dialog, other platforms keep the shared wording', () => {
  it('leaves the descriptions untouched', () => {
    const handle = mount('default');
    handle.open('number');
    expect(document.querySelector('.fc-fmtdlg__hintbar')?.textContent).toBe(
      ja.formatDialog.descGeneral,
    );
  });
});
