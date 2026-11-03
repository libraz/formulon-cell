import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';
import { showRemoveDuplicatesDialog } from '../../../src/toolbar/dialogs/remove-duplicates.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

describe('showRemoveDuplicatesDialog', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('renders an Excel-style column checklist and validates selection', async () => {
    const pending = showRemoveDuplicatesDialog({
      title: 'Remove Duplicates',
      columnsLabel: 'Columns',
      headerLabel: 'My data has headers',
      selectAllLabel: 'Select All',
      unselectAllLabel: 'Unselect All',
      noColumnsLabel: 'Select at least one column.',
      columns: [
        { value: '0', label: 'Column A' },
        { value: '1', label: 'Column B' },
      ],
      initialColumns: ['0', '1'],
      initialHasHeader: true,
      okLabel: 'OK',
      cancelLabel: 'Cancel',
    });

    const dialog = document.querySelector<HTMLElement>('.fc-tb__dlg');
    const list = dialog?.querySelector<HTMLElement>('.fc-dedupedlg__column-list');
    expect(list?.getAttribute('role')).toBe('group');
    expect(list?.getAttribute('aria-label')).toBe('Columns');
    expect(dialog?.querySelectorAll('.fc-dedupedlg__column')).toHaveLength(2);

    dialog?.querySelector<HTMLButtonElement>('.fc-dedupedlg__actions button:nth-child(2)')?.click();
    for (const checkbox of Array.from(
      dialog?.querySelectorAll<HTMLInputElement>('fieldset input') ?? [],
    )) {
      expect(checkbox.checked).toBe(false);
    }
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    expect(dialog?.querySelector<HTMLElement>('.fc-tb__dlg__error')?.textContent).toBe(
      'Select at least one column.',
    );

    dialog?.querySelector<HTMLButtonElement>('.fc-dedupedlg__actions button')?.click();
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(pending).resolves.toEqual({
      columns: ['0', '1'],
      hasHeader: true,
    });
  });

  const macOptions = (host: HTMLElement) => ({
    host,
    title: '重複の削除',
    columnsLabel: '列',
    headerLabel: '先頭行を列見出しとする',
    selectAllLabel: 'すべて選択',
    unselectAllLabel: 'すべて選択解除',
    noColumnsLabel: '1 つ以上の列を選択してください。',
    columns: [
      { value: '0', label: 'A', heading: '名前' },
      { value: '1', label: 'B', heading: '数量' },
      { value: '2', label: 'C', heading: 'メモ' },
    ],
    initialColumns: ['0', '1', '2'],
    initialHasHeader: true,
    okLabel: 'OK',
    cancelLabel: 'キャンセル',
  });
  const macHost = (): HTMLElement => {
    const host = document.createElement('div');
    host.className = 'fc-host';
    host.dataset.fcPlatform = 'mac';
    document.body.appendChild(host);
    return host;
  };

  it('labels the columns by header text, column letters when there is no header row', () => {
    // Excel: with 先頭行を列見出しとする checked the rows read 名前 / 数量 / メモ.
    void showRemoveDuplicatesDialog(macOptions(macHost()));
    const dialog = document.querySelector<HTMLElement>('.fc-tb__dlg');
    const labels = () =>
      Array.from(
        dialog?.querySelectorAll('.fc-dedupedlg__column:not(.fc-dedupedlg__select-all)') ?? [],
        (row) => row.textContent,
      );
    expect(labels()).toEqual(['名前', '数量', 'メモ']);
    const header = dialog?.querySelector<HTMLInputElement>('.fc-dedupedlg__header input');
    header?.click();
    expect(labels()).toEqual(['A', 'B', 'C']);
  });

  it('puts a tri-state すべて選択 row at the top of the Mac list instead of buttons', async () => {
    const pending = showRemoveDuplicatesDialog(macOptions(macHost()));
    const dialog = document.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(dialog?.querySelector('.fc-dedupedlg__actions')).toBeNull();
    const rows = Array.from(dialog?.querySelectorAll('.fc-dedupedlg__column') ?? []);
    expect(rows.map((row) => row.textContent)).toEqual(['すべて選択', '名前', '数量', 'メモ']);
    const all = rows[0]?.querySelector('input') as HTMLInputElement;
    const first = rows[1]?.querySelector('input') as HTMLInputElement;
    expect(all.checked).toBe(true);
    first.click();
    expect([all.checked, all.indeterminate]).toEqual([false, true]);
    all.click();
    expect(
      rows.slice(1).every((row) => (row.querySelector('input') as HTMLInputElement).checked),
    ).toBe(true);
    all.click();
    expect(
      rows.slice(1).some((row) => (row.querySelector('input') as HTMLInputElement).checked),
    ).toBe(false);
    all.click();
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(pending).resolves.toEqual({ columns: ['0', '1', '2'], hasHeader: true });
  });

  it('keeps Remove Duplicates checklist on compact desktop list geometry', () => {
    const css = readFileSync(
      join(root, 'src/styles/core/app/dialog-modules/remove-duplicates.css'),
      'utf8',
    );

    expect(css).toMatch(
      /\.fc-dedupedlg__column-list\s*\{[\s\S]*?gap: 0;[\s\S]*?padding: 0;[\s\S]*?border-radius: 2px;/,
    );
    expect(css).toMatch(
      /\.fc-dedupedlg__column\s*\{[\s\S]*?min-height: 28px;[\s\S]*?padding: 3px 8px;[\s\S]*?border-bottom: 1px solid var\(--fc-rule-subtle/,
    );
    expect(css).toMatch(
      /\.fc-dedupedlg__column:hover,[\s\S]*?\.fc-dedupedlg__column:focus-within\s*\{[\s\S]*?background: var\(--fc-bg-hover/,
    );
    expect(css).not.toContain('background: var(--fc-accent-soft');
  });
});
