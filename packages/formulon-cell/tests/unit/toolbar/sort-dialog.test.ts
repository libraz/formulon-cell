import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';
import { macSortDialogJa } from '../../../src/i18n/strings/mac/sort-dialog.js';
import { showSortDialog } from '../../../src/toolbar/dialogs/sort.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

const labels = {
  title: 'Sort',
  columnLabel: 'Sort by',
  thenByLabel: 'Then by',
  noThenByLabel: '(none)',
  orderLabel: 'Order',
  headerLabel: 'My data has headers',
  addLevelLabel: 'Add Level',
  deleteLevelLabel: 'Delete Level',
  copyLevelLabel: 'Copy Level',
  levelUnavailableLabel: 'At least one sort level is required.',
  ascendingLabel: 'A to Z',
  descendingLabel: 'Z to A',
  okLabel: 'OK',
  cancelLabel: 'Cancel',
};

describe('showSortDialog', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('renders an Excel-style level grid and returns edited levels', async () => {
    const resultPromise = showSortDialog({
      ...labels,
      columns: [
        { value: '0', label: 'Column A' },
        { value: '1', label: 'Column B' },
      ],
      initialColumn: '0',
      initialDirection: 'asc',
      initialHasHeader: true,
    });

    const dialog = document.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(dialog?.textContent).toContain('Add Level');
    expect(dialog?.textContent).toContain('Delete Level');
    expect(dialog?.textContent).toContain('Copy Level');
    expect(dialog?.querySelectorAll('.fc-sortdlg__level')).toHaveLength(1);

    const deleteLevel = dialog?.querySelector<HTMLButtonElement>('.fc-sortdlg__delete-level');
    expect(deleteLevel?.disabled).toBe(true);
    expect(deleteLevel?.dataset.disabledReason).toBe(labels.levelUnavailableLabel);
    expect(deleteLevel?.getAttribute('aria-description')).toBe(labels.levelUnavailableLabel);

    dialog?.querySelector<HTMLButtonElement>('.fc-sortdlg__add-level')?.click();
    expect(dialog?.querySelectorAll('.fc-sortdlg__level')).toHaveLength(2);
    expect(deleteLevel?.disabled).toBe(false);
    expect(deleteLevel?.dataset.disabledReason).toBeUndefined();

    const rows = Array.from(dialog?.querySelectorAll<HTMLElement>('.fc-sortdlg__level') ?? []);
    const secondColumn = rows[1]?.querySelector<HTMLSelectElement>('select[aria-label="Then by"]');
    const secondDirection = rows[1]?.querySelectorAll<HTMLSelectElement>('select')[1];
    if (secondColumn) secondColumn.value = '1';
    if (secondDirection) secondDirection.value = 'desc';

    dialog?.querySelector<HTMLButtonElement>('.fc-sortdlg__copy-level')?.click();
    expect(dialog?.querySelectorAll('.fc-sortdlg__level')).toHaveLength(3);

    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(resultPromise).resolves.toEqual({
      column: '0',
      direction: 'asc',
      levels: [
        { column: '0', direction: 'asc' },
        { column: '1', direction: 'desc' },
        { column: '1', direction: 'desc' },
      ],
      hasHeader: true,
      orientation: 'topToBottom',
      matchCase: false,
    });
  });

  it('focuses the first sort column before the opener command returns', () => {
    const opener = document.createElement('button');
    document.body.appendChild(opener);
    opener.focus();

    const resultPromise = showSortDialog({
      ...labels,
      columns: [{ value: '0', label: 'Column A' }],
      initialColumn: '0',
      initialDirection: 'asc',
      initialHasHeader: false,
    });

    const dialog = document.querySelector<HTMLElement>('.fc-tb__dlg');
    const firstColumn = dialog?.querySelector<HTMLSelectElement>(
      '.fc-sortdlg__level select[aria-label="Sort by"]',
    );
    expect(document.activeElement).toBe(firstColumn);

    dialog?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(document.querySelector('.fc-tb__dlg')).toBeNull();
    expect(document.activeElement).toBe(opener);
    return expect(resultPromise)
      .resolves.toBeNull()
      .finally(() => opener.remove());
  });

  it('keeps Custom Sort levels on compact desktop grid geometry', () => {
    const css = readFileSync(join(root, 'src/styles/core/app/dialog-modules/sort.css'), 'utf8');

    expect(css).toMatch(
      /\.fc-sortdlg__levels\s*\{[\s\S]*?gap: 0;[\s\S]*?padding: 0;[\s\S]*?border-radius: 2px;/,
    );
    expect(css).toMatch(
      /\.fc-sortdlg__level\s*\{[\s\S]*?min-height: 32px;[\s\S]*?padding: 4px 8px;[\s\S]*?border-bottom: 1px solid var\(--fc-rule-subtle/,
    );
    expect(css).toMatch(/\.fc-sortdlg__level--selected\s*\{[\s\S]*?background: var\(--fc-bg-hover/);
    expect(css).toMatch(
      /\.fc-sortdlg__level--selected\s*\{[\s\S]*?box-shadow: inset 0 0 0 1px var\(--fc-fmtdlg-list-focus-border\);/,
    );
    expect(css).not.toContain('background: var(--fc-accent-soft');
  });

  describe('Mac layout', () => {
    const macHost = (): HTMLElement => {
      const host = document.createElement('div');
      host.className = 'fc-host';
      host.dataset.fcPlatform = 'mac';
      document.body.appendChild(host);
      return host;
    };
    const macOptions = (host: HTMLElement) => ({
      ...labels,
      host,
      title: '並べ替え',
      columnLabel: '最優先されるキー',
      thenByLabel: '次に優先されるキー',
      orderLabel: '順序',
      headerLabel: macSortDialogJa.sortDialogHeader ?? '',
      addLevelLabel: 'レベルの追加',
      deleteLevelLabel: 'レベルの削除',
      copyLevelLabel: 'レベルのコピー',
      ascendingLabel: macSortDialogJa.sortDialogAscending ?? '',
      descendingLabel: macSortDialogJa.sortDialogDescending ?? '',
      okLabel: 'OK',
      cancelLabel: 'キャンセル',
      mac: macSortDialogJa,
      columns: [
        { value: '0', label: 'A', heading: '名前' },
        { value: '1', label: 'B', heading: '部門' },
        { value: '2', label: 'C', heading: '売上' },
      ],
      initialColumn: '0',
      initialDirection: 'asc' as const,
      initialHasHeader: true,
    });

    it('lays out the level table and buttons as Excel for Mac does', async () => {
      // Excel ja: 並べ替えるレベルを追加: | 先頭行を列見出しとする on top; columns
      // 列 / 並べ替えのキー / 順序 / 色/アイコン; + / − / コピー, then オプション... / キャンセル / OK.
      const pending = showSortDialog(macOptions(macHost()));
      const dialog = document.querySelector<HTMLElement>('.fc-tb__dlg');
      expect(dialog?.querySelector('.fc-fmtdlg__panel')?.classList.contains('fc-sortdlg')).toBe(
        true,
      );
      expect(dialog?.querySelector('.fc-sortdlg__heading')?.textContent).toBe(
        '並べ替えるレベルを追加:',
      );
      const header = dialog?.querySelector<HTMLLabelElement>('.fc-sortdlg__header');
      expect(header?.textContent?.trim()).toBe('先頭行を列見出しとする');
      expect(header?.querySelector('input')?.checked).toBe(true);
      expect(
        Array.from(
          dialog?.querySelectorAll('.fc-sortdlg__grid-head [role="columnheader"]') ?? [],
          (cell) => cell.textContent,
        ),
      ).toEqual(['', '列', '並べ替えのキー', '順序', '色/アイコン']);

      const row = dialog?.querySelector<HTMLElement>('.fc-sortdlg__level');
      expect(row?.querySelector('.fc-sortdlg__level-label')?.textContent).toBe('最優先されるキー');
      const column = row?.querySelector<HTMLSelectElement>('.fc-sortdlg__column');
      expect(Array.from(column?.options ?? [], (o) => o.text)).toEqual(['名前', '部門', '売上']);
      const sortOn = row?.querySelector<HTMLSelectElement>('.fc-sortdlg__sort-on');
      expect(Array.from(sortOn?.options ?? [], (o) => o.text)).toEqual(['値']);
      const direction = row?.querySelector<HTMLSelectElement>('.fc-sortdlg__direction');
      expect(Array.from(direction?.options ?? [], (o) => o.text)).toEqual([
        '昇順 (A ～ Z)',
        '降順 (Z ～ A)',
      ]);
      expect(row?.querySelector<HTMLSelectElement>('.fc-sortdlg__color')?.disabled).toBe(true);

      const add = dialog?.querySelector<HTMLButtonElement>('.fc-sortdlg__add-level');
      const del = dialog?.querySelector<HTMLButtonElement>('.fc-sortdlg__delete-level');
      const copy = dialog?.querySelector<HTMLButtonElement>('.fc-sortdlg__copy-level');
      expect([add?.textContent, del?.textContent, copy?.textContent]).toEqual(['+', '−', 'コピー']);
      expect([
        add?.getAttribute('aria-label'),
        del?.getAttribute('aria-label'),
        copy?.getAttribute('aria-label'),
      ]).toEqual(['レベルの追加', 'レベルの削除', 'レベルのコピー']);
      const options = dialog?.querySelector<HTMLButtonElement>('.fc-sortdlg__options');
      expect(options?.textContent).toBe('オプション...');
      expect(options?.disabled).toBe(false);
      expect(
        Array.from(
          dialog?.querySelectorAll('.fc-fmtdlg__footer button') ?? [],
          (b) => b.textContent,
        ),
      ).toEqual(['オプション...', 'キャンセル', 'OK']);

      header?.querySelector('input')?.click();
      expect(Array.from(column?.options ?? [], (o) => o.text)).toEqual(['A', 'B', 'C']);

      dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
      await expect(pending).resolves.toEqual({
        column: '0',
        direction: 'asc',
        levels: [{ column: '0', direction: 'asc' }],
        hasHeader: false,
        orientation: 'topToBottom',
        matchCase: false,
      });
    });

    it('keeps the shared layout off the Mac profile', () => {
      void showSortDialog({ ...macOptions(macHost()), host: undefined });
      const dialog = document.querySelector<HTMLElement>('.fc-tb__dlg');
      expect(dialog?.querySelector('.fc-sortdlg__heading')).toBeNull();
      expect(dialog?.querySelector('.fc-sortdlg__sort-on')).toBeNull();
      expect(dialog?.querySelector('.fc-sortdlg__options')).toBeNull();
      expect(dialog?.querySelector('.fc-sortdlg__add-level')?.textContent).toBe('レベルの追加');
      expect(
        Array.from(
          dialog?.querySelector<HTMLSelectElement>('.fc-sortdlg__column')?.options ?? [],
          (o) => o.text,
        ),
      ).toEqual(['A', 'B', 'C']);
    });

    it('pins the measured 700 x 327 Mac frame', () => {
      const css = readFileSync(join(root, 'src/styles/core/app/dialog-modules/sort.css'), 'utf8');
      const mac = '.fc-overlay-portal[data-fc-platform="mac"] .fc-sortdlg';
      const rule = (selector: string): string => {
        const at = css.indexOf(`${selector} {`);
        expect(at, selector).toBeGreaterThan(-1);
        return css.slice(at, css.indexOf('}', at));
      };
      expect(rule(mac)).toMatch(/width: 700px;[\s\S]*height: 327px;/);
      expect(rule(`${mac} .fc-sortdlg__heading`)).toMatch(/top: 43px;[\s\S]*left: 19px;/);
      expect(rule(`${mac} .fc-sortdlg__header`)).toMatch(/top: 43px;[\s\S]*left: 524px;/);
      expect(rule(`${mac} .fc-sortdlg__grid-head`)).toMatch(
        /top: 65px;[\s\S]*left: 20px;[\s\S]*width: 660px;[\s\S]*height: 182px;/,
      );
      expect(css).toContain('grid-template-columns: 111px 167px 99px 153px 110px;');
      expect(rule(`${mac} .fc-sortdlg__levels`)).toMatch(/top: 103px;[\s\S]*left: 20px;/);
      expect(rule(`${mac} .fc-sortdlg__toolbar .fc-fmtdlg__btn`)).toMatch(/top: 254px;/);
      expect(rule(`${mac} .fc-sortdlg__add-level`)).toMatch(/left: 19px;/);
      expect(rule(`${mac} .fc-sortdlg__delete-level`)).toMatch(/left: 42px;/);
      expect(rule(`${mac} .fc-sortdlg__copy-level`)).toMatch(/left: 65px;[\s\S]*width: 49px;/);
      expect(rule(`${mac} .fc-sortdlg__options`)).toMatch(/left: 404px;[\s\S]*width: 97px;/);
    });
  });
});
