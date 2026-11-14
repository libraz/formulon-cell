import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { upsertDefinedName } from '../../../../src/commands/named-ranges.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { scaledTimeout } from '../../../timeouts.js';
import { stubHelpers } from './fixtures.js';

vi.setConfig({ testTimeout: scaledTimeout(20_000) });

const MENU = '#menu-mac-formulas-use-in-formula';

describe('Mac Formulas > Use in Formula menu', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    sheet = await mountStubSheet({
      workbook: await WorkbookHandle.createDefault(),
      locale: 'en',
      ui: { platform: 'mac' },
    });
    sheet.workbook.addSheet('Second');
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  const define = (name: string, localSheetId = -1, attrs: { hidden?: boolean } = {}): void => {
    const result = upsertDefinedName(sheet.workbook, name, '$A$1', localSheetId, attrs);
    expect(result.ok).toBe(true);
  };

  const mountRibbon = () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('formulas');
    return tb;
  };

  const open = (): HTMLElement => {
    host
      .querySelector<HTMLButtonElement>('[data-ribbon-command="mac.formulas.useInFormula"]')
      ?.click();
    const menu = host.querySelector<HTMLElement>(MENU);
    if (!menu) throw new Error('Use in Formula menu was not rendered');
    return menu;
  };

  const rowNames = (menu: HTMLElement): string[] =>
    [...menu.querySelectorAll<HTMLElement>('[data-defined-name-insert]')]
      .filter((el) => el.tagName === 'BUTTON')
      .map((el) => el.textContent ?? '');

  it('lists names sorted case-insensitively, Latin before kana before kanji', () => {
    for (const name of ['beta', '漢字', 'Alpha', 'あいう', '_under', 'alpha2', 'a.b', 'カタ']) {
      define(name);
    }
    mountRibbon();
    const menu = open();
    expect(rowNames(menu)).toEqual([
      '_under',
      'a.b',
      'Alpha',
      'alpha2',
      'beta',
      'あいう',
      'カタ',
      '漢字',
    ]);
  });

  it('lists the active sheet local names bare, and skips other sheets and hidden names', () => {
    define('Book');
    define('Here', 0);
    define('There', 1);
    define('Secret', -1, { hidden: true });
    mountRibbon();
    expect(rowNames(open())).toEqual(['Book', 'Here']);
  });

  it('lists a name shared by a sheet-local and a workbook name once', () => {
    define('Same');
    define('Same', 0);
    mountRibbon();
    expect(rowNames(open())).toEqual(['Same']);
  });

  it('keeps many names in one menu, then a separator and Paste Names last', () => {
    for (let i = 0; i < 44; i += 1) define(`Item_${i}`);
    mountRibbon();
    const menu = open();
    expect(rowNames(menu)).toHaveLength(44);
    const children = [...menu.children];
    const paste = children.at(-1);
    expect(paste?.getAttribute('data-ribbon-command')).toBe('mac.formulas.useInFormula.paste');
    const before = children.at(-2);
    expect(before?.getAttribute('role')).toBe('separator');
    expect(children.slice(0, 44).every((el) => el.getAttribute('role') === 'menuitem')).toBe(true);
    expect(menu.querySelector('[data-more]')).toBeNull();
  });

  it('writes =Name into the active cell when a name is clicked', () => {
    define('Alpha');
    define('Beta');
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 4, col: 2 });
    mountRibbon();
    const menu = open();
    const row = [...menu.querySelectorAll<HTMLButtonElement>('[data-defined-name-insert]')].find(
      (el) => el.textContent === 'Beta',
    );
    row?.click();
    const cell = sheet.instance.store.getState().data.cells.get('0:4:2');
    expect(cell?.formula).toBe('=Beta');
    expect(menu.hidden).toBe(true);
  });

  it('shows a name added after the first open, without duplicate rows', () => {
    define('One');
    mountRibbon();
    expect(rowNames(open())).toEqual(['One']);
    define('Two');
    const menu = open();
    expect(rowNames(menu)).toEqual(['One', 'Two']);
    expect(menu.querySelectorAll('[role="separator"]')).toHaveLength(1);
    expect(
      menu.querySelectorAll('[data-ribbon-command="mac.formulas.useInFormula.paste"]'),
    ).toHaveLength(1);
  });
});
