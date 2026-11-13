import { afterEach, describe, expect, it, vi } from 'vitest';
import { activateSheetView } from '../../../src/commands/sheet-views.js';
import { openSheetViewOptionsDialog } from '../../../src/interact/sheet-view-options-dialog.js';
import { Spreadsheet } from '../../../src/mount.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';
import { stubHelpers } from '../mount/toolbar/fixtures.js';

const dialog = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-sheetviewopts');
const entries = (): [string, boolean][] =>
  [...(dialog()?.querySelectorAll<HTMLElement>('[role="option"]') ?? [])].map((option) => [
    option.textContent ?? '',
    option.getAttribute('aria-selected') === 'true',
  ]);
const button = (root: ParentNode | null, label: string): HTMLButtonElement | undefined =>
  [...(root?.querySelectorAll<HTMLButtonElement>('button') ?? [])].find(
    (candidate) => candidate.textContent === label,
  );
const pick = (name: string): void =>
  [...(dialog()?.querySelectorAll<HTMLElement>('[role="option"]') ?? [])]
    .find((option) => option.textContent === name)
    ?.click();
const prompt = (): HTMLElement | null => document.querySelector<HTMLElement>('.fc-tb__dlg');
const flush = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 0));

describe('Sheet View Options', () => {
  let sheet: MountedStubSheet | undefined;
  afterEach(() => {
    sheet?.dispose();
    sheet = undefined;
    document.body.replaceChildren();
  });

  const mount = async () => {
    sheet = await mountStubSheet({ locale: 'ja', ui: { platform: 'mac' } });
    const { instance } = sheet;
    mutators.upsertSheetView(instance.store, { id: 'v1', name: 'Ops', sheet: 0, hiddenRows: [2] });
    mutators.upsertSheetView(instance.store, { id: 'v2', name: 'Other', sheet: 1 });
    return instance;
  };

  it('lists Default and the sheet’s views with the shown one selected, under the desktop labels', async () => {
    const instance = await mount();
    activateSheetView(instance.store, 'v1');
    openSheetViewOptionsDialog(instance);

    expect(dialog()?.querySelector('.fc-fmtdlg__header')?.textContent).toBe(
      'シート ビューのオプション',
    );
    expect(entries()).toEqual([
      ['既定', false],
      ['Ops', true],
    ]);
    expect([...(dialog()?.querySelectorAll('button') ?? [])].map((b) => b.textContent)).toEqual(
      expect.arrayContaining(['名前の変更', '複製', '削除', 'キャンセル', '切り替え']),
    );
    pick('既定');
    expect(button(dialog(), '削除')?.disabled).toBe(true);
  });

  it('switches to the picked view and closes', async () => {
    const instance = await mount();
    const changed = vi.fn();
    openSheetViewOptionsDialog(instance, { onViewChanged: changed });
    pick('Ops');
    button(dialog(), '切り替え')?.click();

    expect(instance.store.getState().sheetViews.activeViewId).toBe('v1');
    expect([...instance.store.getState().layout.hiddenRows]).toEqual([2]);
    expect(dialog()).toBeNull();
    expect(changed).toHaveBeenCalledOnce();
  });

  it('refuses the desktop’s three rename cases, then renames', async () => {
    const instance = await mount();
    mutators.upsertSheetView(instance.store, { id: 'v3', name: 'Taken', sheet: 0 });
    openSheetViewOptionsDialog(instance);
    const submit = (value: string): string | null | undefined => {
      const input = prompt()?.querySelector<HTMLInputElement>('input[type="text"]');
      if (input) input.value = value;
      prompt()?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
      return prompt()?.querySelector('.fc-tb__dlg__error')?.textContent;
    };

    pick('既定');
    button(dialog(), '名前の変更')?.click();
    expect(submit('Main')).toBe(
      '他のシート ビューの名前は変更できますが、既定のビューは変更できません。',
    );
    button(prompt(), 'キャンセル')?.click();
    await flush();

    pick('Ops');
    button(dialog(), '名前の変更')?.click();
    expect(prompt()?.querySelector<HTMLInputElement>('input[type="text"]')?.value).toBe('Ops');
    expect(submit('a;b')).toBe('空ではなく、かつ ";" が含まれていない名前を入力してください');
    expect(submit('taken')).toBe(
      '同じ名前のシート ビューが既に存在します。別の名前を入力するか、もう一方のシート ビューを削除してから、もう一度お試しください。',
    );
    submit('Review');
    await flush();
    expect(instance.store.getState().sheetViews.views.find((v) => v.id === 'v1')?.name).toBe(
      'Review',
    );
    expect(entries().map(([name]) => name)).toEqual(['既定', 'Review', 'Taken']);
  });

  it('duplicates a view under its name with the next number', async () => {
    const instance = await mount();
    openSheetViewOptionsDialog(instance);
    pick('Ops');
    button(dialog(), '複製')?.click();

    const copy = instance.store.getState().sheetViews.views.find((v) => v.name === 'Ops 2');
    expect(copy).toMatchObject({ sheet: 0, hiddenRows: [2] });
    expect(copy?.id).not.toBe('v1');
    expect(entries().map(([name]) => name)).toEqual(['既定', 'Ops', 'Ops 2']);
  });

  it('asks before deleting, naming the view', async () => {
    const instance = await mount();
    openSheetViewOptionsDialog(instance);
    pick('Ops');
    button(dialog(), '削除')?.click();

    expect(prompt()?.querySelector('.fc-tb__dlg__headline')?.textContent).toBe(
      'このシート ビューを削除しますか?',
    );
    expect(prompt()?.querySelector('.fc-tb__dlg__message')?.textContent).toBe(
      '"Ops" を削除しますか?',
    );
    button(prompt(), '削除')?.click();
    await flush();
    expect(instance.store.getState().sheetViews.views.map((v) => v.id)).toEqual(['v2']);
    expect(entries()).toEqual([['既定', true]]);
  });

  it('opens from the ribbon’s Options button', async () => {
    const instance = await mount();
    const host = document.createElement('div');
    document.body.appendChild(host);
    const tb = Spreadsheet.mountToolbar(host, instance, { helpers: stubHelpers() });
    expect(tb.applyCommand('mac.view.sheetViewOptions')).toBe(true);
    expect(dialog()).not.toBeNull();
    tb.dispose();
  });
});
