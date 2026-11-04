import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';
import { showProtectSheetDialog } from '../../../src/toolbar/dialogs/protection.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const read = (path: string): string => readFileSync(join(root, path), 'utf8');

const strings = {
  title: 'シートの保護',
  password: 'パスワード (省略可):',
  passwordPlaceholder: '',
  confirmPassword: '確認:',
  passwordMismatch: 'x',
  allowLabel: 'このシートのすべてのユーザーに許可する操作:',
  allowSelectLockedCells: 'ロックされたセルの選択',
  allowSelectUnlockedCells: 'ロックされていないセルの選択',
  allowFormatCells: 'セルの書式設定',
  allowFormatColumns: '列の書式設定',
  allowFormatRows: '行の書式設定',
  allowInsertColumns: '列の挿入',
  allowInsertRows: '行の挿入',
  allowInsertHyperlinks: 'ハイパーリンクの挿入',
  allowDeleteColumns: '列の削除',
  allowDeleteRows: '行の削除',
  allowSort: '並べ替え',
  allowAutoFilter: 'オートフィルターの使用',
  allowPivotTables: 'ピボットテーブルとピボットグラフを使用する',
  allowEditObjects: 'オブジェクトの編集',
  allowEditScenarios: 'シナリオの編集',
  ok: 'OK',
  cancel: 'キャンセル',
};

describe('Protect Sheet permission group', () => {
  afterEach(() => document.body.replaceChildren());

  it('fills the Mac group box column by column, eight options on the left', () => {
    void showProtectSheetDialog({ strings });
    const group = document.querySelector<HTMLElement>('.fc-protectdlg__choices');
    expect(group?.getAttribute('role')).toBe('group');
    expect(group?.getAttribute('aria-label')).toBe(strings.allowLabel);
    const labels = [...(group?.querySelectorAll('label') ?? [])].map((l) => l.textContent);
    expect(labels.slice(0, 8)).toEqual([
      'ロックされたセルの選択',
      'ロックされていないセルの選択',
      'セルの書式設定',
      '列の書式設定',
      '行の書式設定',
      '列の挿入',
      '行の挿入',
      'ハイパーリンクの挿入',
    ]);
    expect(labels[8]).toBe('列の削除');

    expect(read('src/styles/core/app.css')).toContain(
      '@import "./app/dialog-modules/protect-sheet.css";',
    );
    const css = read('src/styles/core/app/dialog-modules/protect-sheet.css');
    const rule = css.slice(
      css.indexOf('.fc-overlay-portal[data-fc-platform="mac"] .fc-protectdlg__choices {'),
    );
    const body = rule.slice(0, rule.indexOf('}'));
    expect(body).toContain('grid-auto-flow: column;');
    expect(body).toContain('grid-template-rows: repeat(8, auto);');
    expect(body).toContain('grid-template-columns: repeat(2, minmax(0, 1fr));');
  });
});
