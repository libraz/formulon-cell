import { afterEach, describe, expect, it } from 'vitest';
import { macSortDialogJa } from '../../../src/i18n/strings/mac/sort-dialog.js';
import { showSortDialog } from '../../../src/toolbar/dialogs/sort.js';

const macOptions = (host: HTMLElement) => ({
  host,
  title: '並べ替え',
  columnLabel: '最優先されるキー',
  thenByLabel: '次に優先されるキー',
  noThenByLabel: '(なし)',
  orderLabel: '順序',
  headerLabel: macSortDialogJa.sortDialogHeader ?? '',
  addLevelLabel: 'レベルの追加',
  deleteLevelLabel: 'レベルの削除',
  copyLevelLabel: 'レベルのコピー',
  levelUnavailableLabel: '-',
  ascendingLabel: '昇順',
  descendingLabel: '降順',
  okLabel: 'OK',
  cancelLabel: 'キャンセル',
  mac: macSortDialogJa,
  columns: [
    { value: '0', label: 'A', heading: '名前' },
    { value: '1', label: 'B', heading: '部門' },
  ],
  rows: [
    { value: '0', label: '行 1' },
    { value: '1', label: '行 2' },
  ],
  initialColumn: '0',
  initialDirection: 'asc' as const,
  initialHasHeader: true,
});

const open = () => {
  const host = document.createElement('div');
  host.className = 'fc-host';
  host.dataset.fcPlatform = 'mac';
  document.body.appendChild(host);
  const pending = showSortDialog(macOptions(host));
  const dialog = document
    .querySelector<HTMLElement>('.fc-sortdlg')
    ?.closest<HTMLElement>('.fc-tb__dlg');
  const optionsBtn = dialog?.querySelector<HTMLButtonElement>('.fc-sortdlg__options');
  return { pending, dialog, optionsBtn };
};

const sheet = () => document.querySelector<HTMLElement>('.fc-sortdlg-options');
const radios = (name: string) =>
  Array.from(
    sheet()?.querySelectorAll<HTMLInputElement>(`input[type="radio"][name="${name}"]`) ?? [],
  );
const button = (text: string) =>
  Array.from(sheet()?.querySelectorAll<HTMLButtonElement>('button') ?? []).find(
    (b) => b.textContent === text,
  );

describe('Mac sort options sheet', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('enables the options button and opens a sheet with measured controls and defaults', () => {
    const { optionsBtn } = open();
    expect(optionsBtn?.disabled).toBe(false);
    optionsBtn?.click();
    expect(sheet()).not.toBeNull();
    const text = (el: HTMLInputElement) => el.parentElement?.textContent?.trim();
    const orient = radios('fc-sort-orientation');
    expect(orient.map(text)).toEqual(['行単位', '列単位']);
    expect(orient.map((r) => r.checked)).toEqual([true, false]);
    const method = radios('fc-sort-method');
    expect(method.map(text)).toEqual(['ふりがなを使う', 'ふりがなを使わない']);
    expect(method.map((r) => r.checked)).toEqual([true, false]);
    const matchCase = sheet()?.querySelector<HTMLInputElement>('input[type="checkbox"]');
    expect(matchCase?.parentElement?.textContent?.trim()).toBe('大文字と小文字を区別する');
    expect(matchCase?.checked).toBe(false);
    expect(button('キャンセル')).toBeDefined();
    expect(button('OK')).toBeDefined();
  });

  it('returns orientation and matchCase after OK in the sheet, and switches the list to rows', async () => {
    const { pending, dialog, optionsBtn } = open();
    optionsBtn?.click();
    radios('fc-sort-orientation')[1]?.click();
    sheet()?.querySelector<HTMLInputElement>('input[type="checkbox"]')?.click();
    button('OK')?.click();
    expect(sheet()).toBeNull();
    await Promise.resolve();
    const select = dialog?.querySelector<HTMLSelectElement>('.fc-sortdlg__column');
    expect(Array.from(select?.options ?? [], (o) => o.text)).toEqual(['行 1', '行 2']);
    Array.from(dialog?.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__footer button') ?? [])
      .find((b) => b.textContent === 'OK')
      ?.click();
    const result = await pending;
    expect(result?.orientation).toBe('leftToRight');
    expect(result?.matchCase).toBe(true);
  });

  it('discards the sheet choices on Cancel', async () => {
    const { pending, dialog, optionsBtn } = open();
    optionsBtn?.click();
    radios('fc-sort-orientation')[1]?.click();
    button('キャンセル')?.click();
    optionsBtn?.click();
    expect(radios('fc-sort-orientation').map((r) => r.checked)).toEqual([true, false]);
    button('キャンセル')?.click();
    Array.from(dialog?.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__footer button') ?? [])
      .find((b) => b.textContent === 'OK')
      ?.click();
    const result = await pending;
    expect(result?.orientation).toBe('topToBottom');
    expect(result?.matchCase).toBe(false);
  });
});
