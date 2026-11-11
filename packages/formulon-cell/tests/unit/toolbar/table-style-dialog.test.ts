import { afterEach, describe, expect, it, vi } from 'vitest';
import { en, ja } from '../../../src/i18n/strings.js';
import {
  showTableStyleDialog,
  type TableStyleDialogOptions,
} from '../../../src/toolbar/dialogs/table-style.js';

const dialog = (): HTMLElement | null => document.body.querySelector<HTMLElement>('.fc-tb__dlg');
const options = (): HTMLElement[] =>
  Array.from(dialog()?.querySelectorAll<HTMLElement>('[role="option"]') ?? []);
const button = (label: string): HTMLButtonElement => {
  const found = Array.from(dialog()?.querySelectorAll<HTMLButtonElement>('button') ?? []).find(
    (b) => b.textContent === label,
  );
  if (!found) throw new Error(`no ${label}`);
  return found;
};
const ok = (): void =>
  dialog()?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

const open = (over: Partial<TableStyleDialogOptions> = {}) =>
  showTableStyleDialog({
    kind: 'table',
    mode: 'new',
    mac: true,
    strings: ja,
    initial: { name: 'テーブル スタイル 1', elements: {}, stripeSizes: {} },
    editFormat: vi.fn(),
    describe: () => '',
    ...over,
  });

describe('New Table Style dialog', () => {
  afterEach(() => document.body.replaceChildren());

  it('shows the Mac dialog: its labels and the table elements in the desktop order', async () => {
    const pending = open();
    const text = dialog()?.textContent ?? '';
    expect(
      dialog()?.querySelector('.fc-fmtdlg__header, .fc-tb__dlg-title')?.textContent ?? text,
    ).toContain('新しいテーブル スタイル');
    for (const label of ['名前:', 'テーブル要素:', '書式', '消去', 'プレビュー', '縞のサイズ']) {
      expect(text).toContain(label);
    }
    expect(text).toContain('このブックの既定の表スタイルに設定する');
    expect(
      options()
        .map((o) => o.textContent)
        .slice(0, 7),
    ).toEqual([
      'テーブル全体',
      '最初の列のストライプ',
      '2 番目の列のストライプ',
      '最初の行のストライプ',
      '2 番目の行のストライプ',
      '最後の列',
      '最初の列',
    ]);
    expect(options()).toHaveLength(13);
    // Clear is unavailable until the selected element has a format.
    expect(button('消去').disabled).toBe(true);
    button('キャンセル').click();
    await expect(pending).resolves.toBeNull();
  });

  it('formats and clears elements, sets a stripe size, and returns the style', async () => {
    const editFormat = vi.fn<TableStyleDialogOptions['editFormat']>((_initial, apply) =>
      apply({ fill: '#336699', bold: true }),
    );
    const pending = open({ editFormat, mac: false, strings: en });
    options()
      .find((o) => o.dataset.element === 'headerRow')
      ?.click();
    button('Format').click();
    expect(editFormat).toHaveBeenCalledWith({}, expect.any(Function));
    expect(button('Clear').disabled).toBe(false);
    options()
      .find((o) => o.dataset.element === 'firstRowStripe')
      ?.click();
    button('Format').click();
    const stripe = dialog()?.querySelector<HTMLSelectElement>('[data-dialog-field="stripeSize"]');
    if (!stripe) throw new Error('no stripe size');
    expect(stripe.disabled).toBe(false);
    stripe.value = '2';
    stripe.dispatchEvent(new Event('change'));
    button('Clear').click();
    dialog()?.querySelector<HTMLInputElement>('input[type="checkbox"]')?.click();
    ok();
    await expect(pending).resolves.toEqual({
      name: 'テーブル スタイル 1',
      elements: { headerRow: { fill: '#336699', bold: true } },
      stripeSizes: { firstRowStripe: 2 },
      setAsDefault: true,
    });
  });

  it('lists the PivotTable elements and drops the default checkbox on the Mac', async () => {
    const pending = open({ kind: 'pivot' });
    expect(options()).toHaveLength(25);
    expect(options().at(-1)?.textContent).toBe(ja.ribbonMenu.tableStyleElements.grandTotalRow);
    expect(dialog()?.textContent).toContain('新しいピボットテーブル クイック スタイル');
    expect(dialog()?.textContent).not.toContain(ja.ribbonMenu.tableStyleSetDefaultPivot);
    ok();
    await expect(pending).resolves.toMatchObject({ setAsDefault: false });
  });
});
