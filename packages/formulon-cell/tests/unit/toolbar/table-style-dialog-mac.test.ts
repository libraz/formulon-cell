import { afterEach, describe, expect, it, vi } from 'vitest';
import { ja } from '../../../src/i18n/strings.js';
import {
  showTableStyleDialog,
  type TableStyleDialogOptions,
} from '../../../src/toolbar/dialogs/table-style.js';

const dialog = (): HTMLElement | null => document.body.querySelector<HTMLElement>('.fc-tb__dlg');
const open = (over: Partial<TableStyleDialogOptions>) =>
  showTableStyleDialog({
    kind: 'table',
    mode: 'new',
    mac: true,
    strings: ja,
    initial: { name: 'MyStyle', elements: {}, stripeSizes: {} },
    editFormat: vi.fn(),
    describe: () => '',
    ...over,
  });

// The table and PivotTable style dialog on the Mac desktop (ja).
describe('table style dialog on the Mac profile', () => {
  afterEach(() => document.body.replaceChildren());

  it('names the style in the Modify title, and Duplicate uses that title', () => {
    void open({ mode: 'modify' });
    expect(dialog()?.textContent).toContain('表のスタイルの変更 MyStyle');
    document.body.replaceChildren();
    void open({
      mode: 'duplicate',
      initial: { name: 'TableStyleLight20 2', elements: {}, stripeSizes: {} },
    });
    expect(dialog()?.textContent).toContain('表のスタイルの変更 TableStyleLight20 2');
  });

  it('offers the default checkbox for a PivotTable style, with the table wording', async () => {
    const pending = open({ kind: 'pivot', mode: 'new' });
    const box = dialog()?.querySelector<HTMLInputElement>(
      'input[data-dialog-field="setAsDefault"]',
    );
    expect(box?.closest('label')?.textContent).toBe('このブックの既定の表スタイルに設定する');
    if (box) box.checked = true;
    dialog()?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(pending).resolves.toMatchObject({ setAsDefault: true });
  });

  it('lists the PivotTable elements in the desktop order', () => {
    void open({ kind: 'pivot', mode: 'new' });
    expect(
      [...(dialog()?.querySelectorAll('[role="option"]') ?? [])].map((o) => o.textContent),
    ).toEqual([
      'テーブル全体',
      'レポート フィルター ラベル',
      'レポート フィルターの値',
      '最初の列のストライプ',
      '2 番目の列のストライプ',
      '最初の行のストライプ',
      '2 番目の行のストライプ',
      '最初の列',
      '見出し行',
      '最初の見出しセル',
      '小計列 1',
      '小計列 2',
      '小計列 3',
      '空白行',
      '小計行 1',
      '小計行 2',
      '小計行 3',
      '列小見出し 1',
      '列小見出し 2',
      '列小見出し 3',
      '行小見出し 1',
      '行小見出し 2',
      '行小見出し 3',
      '総計列',
      '総計行',
    ]);
  });
});
