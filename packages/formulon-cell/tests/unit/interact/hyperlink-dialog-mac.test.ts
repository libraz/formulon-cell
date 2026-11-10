import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings.js';
import { attachHyperlinkDialog } from '../../../src/interact/hyperlink-dialog.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const q = <T extends HTMLElement>(sel: string): T => {
  const el = document.querySelector<T>(sel);
  if (!el) throw new Error(`missing ${sel}`);
  return el;
};
const visible = (el: Element | null): boolean => !!el && !(el as HTMLElement).closest('[hidden]');
const type = (input: HTMLInputElement, value: string): void => {
  input.value = value;
  input.dispatchEvent(new Event('input', { bubbles: true }));
};

describe('Insert Hyperlink on the Mac host', () => {
  let host: HTMLElement;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    wb = await WorkbookHandle.createDefault({ preferStub: true });
    wb.addSheet('データ');
    host = document.createElement('div');
    host.className = 'fc-host';
    host.dataset.fcPlatform = 'mac';
    document.body.appendChild(host);
  });
  afterEach(() => {
    wb.dispose();
    document.body.innerHTML = '';
  });

  const open = (seed?: (store: ReturnType<typeof createSpreadsheetStore>) => void) => {
    const store = createSpreadsheetStore();
    seed?.(store);
    const handle = attachHyperlinkDialog({ host, store, getWb: () => wb, strings: ja });
    handle.open();
    return { store, handle };
  };
  const ok = (): HTMLButtonElement => q('.fc-hldlg .fc-fmtdlg__btn--primary');
  const link = (store: ReturnType<typeof createSpreadsheetStore>) =>
    store.getState().format.formats.get(addrKey(store.getState().selection.active));

  it('shows the display text row, the three link tabs and the web page tab', () => {
    const { handle } = open((store) => {
      store.setState((s) => {
        const cells = new Map(s.data.cells);
        cells.set(addrKey({ sheet: 0, row: 0, col: 0 }), {
          value: { kind: 'text', value: '地域' },
          formula: null,
        });
        return { ...s, data: { ...s.data, cells } };
      });
    });
    expect(q('.fc-hldlg__display > span').textContent).toBe('表示するテキスト:');
    expect(q<HTMLInputElement>('.fc-hldlg__display input').value).toBe('地域');
    expect(q('.fc-hldlg__tip').textContent).toBe('ヒント設定...');
    expect(
      [...document.querySelectorAll('.fc-hldlg [role="tab"]')].map((t) => t.textContent),
    ).toEqual(['Web ページまたはファイル', 'このドキュメント', '電子メール アドレス']);
    expect(q('.fc-hldlg__intro').textContent).toBe(
      '既存のファイルまたは Web ページにリンクします。',
    );
    expect(q('.fc-hldlg__web .fc-hldlg__field > span').textContent).toBe('アドレス:');
    expect(q('.fc-hldlg__browse').textContent).toBe('選択...');
    expect(ok().disabled).toBe(true);
    handle.detach();
  });

  it('disables the display text for a multi-cell selection', () => {
    const { handle } = open((store) =>
      mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 6, c1: 3 }),
    );
    const display = q<HTMLInputElement>('.fc-hldlg__display input');
    expect(display.disabled).toBe(true);
    expect(display.value).toBe('<<ドキュメント内の選択範囲>>');
    handle.detach();
  });

  it('links to a place in this document', () => {
    const { store, handle } = open();
    q<HTMLElement>('.fc-hldlg [role="tab"][data-tab="document"]').click();
    expect(visible(document.querySelector('.fc-hldlg__document'))).toBe(true);
    expect(visible(document.querySelector('.fc-hldlg__web'))).toBe(false);
    q<HTMLElement>('.fc-hldlg__places [data-sheet="データ"]').click();
    type(q<HTMLInputElement>('.fc-hldlg__document input'), 'B2');
    expect(ok().disabled).toBe(false);
    ok().click();
    expect(link(store)?.hyperlink).toBe("#'データ'!B2");
    handle.detach();
  });

  it('links to an e-mail address with a subject and remembers it', () => {
    const { store, handle } = open();
    q<HTMLElement>('.fc-hldlg [role="tab"][data-tab="email"]').click();
    const [address, subject] = [
      ...document.querySelectorAll<HTMLInputElement>('.fc-hldlg__email input'),
    ];
    if (!address || !subject) throw new Error('missing e-mail fields');
    type(address, 'a@example.com');
    type(subject, '売上 報告');
    ok().click();
    expect(link(store)?.hyperlink).toBe(
      `mailto:a@example.com?subject=${encodeURIComponent('売上 報告')}`,
    );
    handle.open();
    q<HTMLElement>('.fc-hldlg [role="tab"][data-tab="email"]').click();
    expect(
      [...document.querySelectorAll('.fc-hldlg__recent [role="option"]')].map((o) => o.textContent),
    ).toEqual(['a@example.com']);
    handle.detach();
  });

  it('opens an existing in-document link on its tab', () => {
    const { handle } = open((store) =>
      mutators.setRangeFormat(
        store,
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        { hyperlink: '#Sheet1!C3' },
      ),
    );
    expect(q('.fc-hldlg [role="tab"][aria-selected="true"]').dataset.tab).toBe('document');
    expect(q<HTMLInputElement>('.fc-hldlg__document input').value).toBe('C3');
    handle.detach();
  });
});
