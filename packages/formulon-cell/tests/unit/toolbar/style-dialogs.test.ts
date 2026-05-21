import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';

import { ja } from '../../../src/i18n/strings.js';
import { showCellStyleDialog } from '../../../src/toolbar/dialogs/cell-style.js';
import { showTableStyleDialog } from '../../../src/toolbar/dialogs/table-style.js';

const sourcePath = (relative: string): string =>
  join(process.cwd(), 'src/toolbar/dialogs', relative);

const closeDialog = async (): Promise<void> => {
  document.body
    .querySelector<HTMLButtonElement>('.fc-fmtdlg__btn:not(.fc-fmtdlg__btn--primary)')
    ?.click();
  await Promise.resolve();
};

const nextFrame = (): Promise<void> =>
  new Promise((resolve) => {
    requestAnimationFrame(() => resolve());
  });

describe('style dialogs i18n', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('renders Create Table Style labels from the provided dictionary', async () => {
    const pending = showTableStyleDialog({
      title: ja.ribbonMenu.tableStyleNew,
      strings: ja,
    });

    const dialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(dialog?.textContent).toContain('新しい表スタイル');
    expect(dialog?.textContent).toContain('スタイル名');
    expect(dialog?.textContent).toContain('スタイルの種類');
    expect(dialog?.textContent).toContain('アクセント色');
    expect(dialog?.textContent).toContain('縞模様の行');
    expect(dialog?.textContent).toContain('最初の列を強調');
    expect(dialog?.textContent).not.toContain('Style name');
    expect(dialog?.textContent).not.toContain('First column emphasis');

    await closeDialog();
    await expect(pending).resolves.toBeNull();
  });

  it('focuses and selects the table style name field on open', async () => {
    const pending = showTableStyleDialog({
      title: ja.ribbonMenu.tableStyleNew,
      strings: ja,
    });
    await nextFrame();

    const name = document.body.querySelector<HTMLInputElement>('input[data-dialog-field="name"]');
    expect(document.activeElement).toBe(name);
    expect(name?.selectionStart).toBe(0);
    expect(name?.selectionEnd).toBe(name?.value.length);

    await closeDialog();
    await expect(pending).resolves.toBeNull();
  });

  it('renders Create Cell Style labels from the provided dictionary', async () => {
    const pending = showCellStyleDialog({
      title: ja.ribbonMenu.cellStyleNew,
      strings: ja,
    });

    const dialog = document.body.querySelector<HTMLElement>('.app__dlg');
    expect(dialog?.textContent).toContain('新しいセルのスタイル');
    expect(dialog?.textContent).toContain('スタイル名');
    expect(dialog?.textContent).toContain('スタイルに含めるもの');
    expect(dialog?.textContent).toContain('表示形式');
    expect(dialog?.textContent).toContain('配置');
    expect(dialog?.textContent).toContain('フォント');
    expect(dialog?.textContent).toContain('罫線');
    expect(dialog?.textContent).toContain('塗りつぶし');
    expect(dialog?.textContent).toContain('保護');
    expect(dialog?.textContent).not.toContain('Style includes');
    expect(dialog?.textContent).not.toContain('Protection');

    await closeDialog();
    await expect(pending).resolves.toBeNull();
  });

  it('focuses and selects the cell style name field on open', async () => {
    const pending = showCellStyleDialog({
      title: ja.ribbonMenu.cellStyleNew,
      strings: ja,
    });
    await nextFrame();

    const name = document.body.querySelector<HTMLInputElement>('input[data-dialog-field="name"]');
    expect(document.activeElement).toBe(name);
    expect(name?.selectionStart).toBe(0);
    expect(name?.selectionEnd).toBe(name?.value.length);

    await closeDialog();
    await expect(pending).resolves.toBeNull();
  });

  it('keeps style dialog name handling on the shared shell helper', () => {
    const sources = [
      readFileSync(sourcePath('table-style.ts'), 'utf8'),
      readFileSync(sourcePath('cell-style.ts'), 'utf8'),
    ];

    for (const source of sources) {
      expect(source).toContain('appendDialogNameField(');
      expect(source).not.toContain("dataset.dialogField = 'name'");
      expect(source).not.toContain('showInputError(');
    }
  });
});
