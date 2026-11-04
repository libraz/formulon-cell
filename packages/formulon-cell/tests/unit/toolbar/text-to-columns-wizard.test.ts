import { afterEach, describe, expect, it } from 'vitest';
import {
  showTextToColumnsDialog,
  type TextToColumnsWizardStrings,
} from '../../../src/toolbar/dialogs/text-to-columns.js';

const strings = {
  title: '区切り位置指定ウィザード',
  dataType: '元のデータの形式',
  delimited: '区切り記号付き',
  fixedWidth: '固定幅',
  fixedWidthUnavailable: '固定幅での分割はまだ使用できません。',
  delimiters: '区切り文字',
  tab: 'タブ',
  semicolon: 'セミコロン',
  comma: 'カンマ',
  space: 'スペース',
  other: 'その他',
  treatConsecutive: '連続した区切り文字は 1 文字として扱う',
  preview: 'データのプレビュー',
  noDelimited: '区切り文字を含むテキストが見つかりません',
  ok: 'OK',
  cancel: 'キャンセル',
};

// Excel for Mac (ja) wizard wording.
const wizard: TextToColumnsWizardStrings = {
  title: '区切り位置指定ウィザード - {step} / 3',
  step1Intro: '選択したデータは区切り文字で区切られています。',
  step1Prompt: '[次へ] をクリックするか、区切るデータの形式を指定してください。',
  delimitedHint: '- コンマやタブなどの区切り文字によってフィールドごとに区切られたデータ',
  fixedWidth: '固定長',
  fixedHint: '- スペースによって右または左に揃えられた固定長フィールドのデータ',
  preview: '選択したデータのプレビュー:',
  step2Prompt: 'フィールドの区切り文字を指定してください。',
  delimiters: '区切り文字',
  other: 'その他:',
  qualifier: '文字列の引用符:',
  qualifierNone: 'なし',
  step3Prompt: '区切ったあとの列のデータ形式を選択してください。',
  columnFormat: '列のデータ形式',
  general: '標準',
  text: '文字列',
  date: '日付:',
  skip: '削除する',
  destination: '表示先:',
  details: '詳細...',
  back: '< 戻る',
  next: '次へ >',
  finish: '完了',
};

const open = (extra: Record<string, unknown> = {}) =>
  showTextToColumnsDialog({ strings, wizard, previewRows: ['a,b,c', 'd,e,f'], ...extra });
const q = <T extends HTMLElement>(sel: string): T => {
  const el = document.querySelector<T>(sel);
  if (!el) throw new Error(`missing ${sel}`);
  return el;
};
const button = (label: string): HTMLButtonElement => {
  const b = Array.from(
    document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__footer button'),
  ).find((el) => el.textContent === label);
  if (!b) throw new Error(`missing button ${label}`);
  return b;
};
const step = (n: number): HTMLElement => q(`.fc-textcols__step[data-step="${n}"]`);

describe('Text to Columns wizard', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('opens on step 1 with the type choices and a disabled Back button', () => {
    const done = open();
    expect(q('.fc-fmtdlg__header').textContent).toBe('区切り位置指定ウィザード - 1 / 3');
    expect(step(1).hidden).toBe(false);
    expect(step(2).hidden).toBe(true);
    expect(step(1).textContent).toContain('選択したデータは区切り文字で区切られています。');
    expect(step(1).textContent).toContain('区切り記号付き');
    expect(step(1).textContent).toContain('固定長');
    expect(
      Array.from(document.querySelectorAll('.fc-fmtdlg__footer button'), (b) => b.textContent),
    ).toEqual(['キャンセル', '< 戻る', '次へ >', '完了']);
    expect(button('< 戻る').disabled).toBe(true);
    button('キャンセル').click();
    return done;
  });

  it('steps through delimiters and data formats, and Back returns', async () => {
    const done = open();
    button('次へ >').click();
    expect(q('.fc-fmtdlg__header').textContent).toBe('区切り位置指定ウィザード - 2 / 3');
    expect(button('< 戻る').disabled).toBe(false);
    const tab = step(2).querySelector<HTMLInputElement>('[data-dialog-field="delimiter-tab"]');
    expect(tab?.checked).toBe(true);
    expect(step(2).textContent).toContain('文字列の引用符:');
    button('次へ >').click();
    expect(q('.fc-fmtdlg__header').textContent).toBe('区切り位置指定ウィザード - 3 / 3');
    expect(button('次へ >').disabled).toBe(true);
    expect(step(3).textContent).toContain('列のデータ形式');
    button('< 戻る').click();
    expect(q('.fc-fmtdlg__header').textContent).toBe('区切り位置指定ウィザード - 2 / 3');
    button('キャンセル').click();
    expect(await done).toBeNull();
  });

  it('finishes from any step with the chosen delimiters and qualifier', async () => {
    const done = open();
    button('次へ >').click();
    const tab = step(2).querySelector<HTMLInputElement>('[data-dialog-field="delimiter-tab"]');
    const comma = step(2).querySelector<HTMLInputElement>('[data-dialog-field="delimiter-,"]');
    if (!tab || !comma) throw new Error('missing delimiter checks');
    tab.checked = false;
    comma.checked = true;
    comma.dispatchEvent(new Event('change', { bubbles: true }));
    const quote = step(2).querySelector<HTMLSelectElement>('select');
    if (!quote) throw new Error('missing qualifier');
    expect(Array.from(quote.options, (o) => o.text)).toContain('{なし}');
    quote.value = "'";
    button('完了').click();
    expect(await done).toEqual({
      delimiters: [','],
      collapseConsecutiveDelimiters: false,
      textQualifier: "'",
    });
  });

  it('keeps the single-page dialog when no wizard strings are given', async () => {
    const done = showTextToColumnsDialog({ strings, previewRows: [] });
    expect(document.querySelector('.fc-textcols__step')).toBeNull();
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__footer button')?.click();
    await done;
  });

  it('splits at fixed column breaks set by clicking the data', async () => {
    const done = open({ previewRows: ['abcdef', 'ghijkl'] });
    const fixed = step(1).querySelectorAll<HTMLInputElement>('input[name="fc-textcols-type"]')[1];
    expect(fixed?.disabled).toBe(false);
    if (!fixed) throw new Error('missing fixed radio');
    fixed.checked = true;
    fixed.dispatchEvent(new Event('change', { bubbles: true }));
    button('次へ >').click();
    expect(step(2).querySelector<HTMLElement>('.fc-textcols__fixed')?.hidden).toBe(false);
    expect(step(2).querySelector<HTMLElement>('.fc-textcols__delimiter-grid')?.hidden).toBe(true);
    const chars = Array.from(
      step(2).querySelectorAll<HTMLElement>('.fc-textcols__fixed-row')[0]?.children ?? [],
    ) as HTMLElement[];
    chars[2]?.click();
    chars[4]?.click();
    chars[4]?.click();
    expect(q('.fc-textcols__preview pre').textContent).toBe('ab | cdef\ngh | ijkl');
    button('完了').click();
    expect(await done).toEqual({
      delimiters: [],
      collapseConsecutiveDelimiters: false,
      textQualifier: null,
      fixedBreaks: [2],
    });
  });

  it('keeps the wizard open when no fixed break is set', () => {
    const done = open({ previewRows: ['abcdef'] });
    const fixed = step(1).querySelectorAll<HTMLInputElement>('input[name="fc-textcols-type"]')[1];
    if (!fixed) throw new Error('missing fixed radio');
    fixed.checked = true;
    fixed.dispatchEvent(new Event('change', { bubbles: true }));
    button('完了').click();
    expect(q('.fc-fmtdlg__header').textContent).toBe('区切り位置指定ウィザード - 2 / 3');
    button('キャンセル').click();
    return done;
  });
});
