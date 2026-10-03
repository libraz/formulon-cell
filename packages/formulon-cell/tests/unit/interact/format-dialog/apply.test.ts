import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../../src/commands/history.js';
import { setCellLocked, setProtectedSheet } from '../../../../src/commands/protection.js';
import { addrKey } from '../../../../src/engine/workbook-handle.js';
import { attachFormatDialog } from '../../../../src/interact/format-dialog.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { mergeWorkbook, setActive, setRange } from './fixtures.js';

describe('attachFormatDialog', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    setActive(store, 0, 0);
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('preserves hyperlink display and tooltip metadata when applying Format Cells', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        hyperlink: 'https://example.test',
        hyperlinkDisplay: 'Example',
        hyperlinkTooltip: 'Open example',
      },
    );
    const handle = attachFormatDialog({ host, store });
    try {
      handle.open('more');
      const ok = Array.from(
        document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__footer button'),
      ).find((b) => b.textContent === 'OK');
      ok?.click();

      const format = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
      expect(format).toMatchObject({
        hyperlink: 'https://example.test',
        hyperlinkDisplay: 'Example',
        hyperlinkTooltip: 'Open example',
      });
    } finally {
      handle.detach();
    }
  });

  it('drops stale hyperlink display and tooltip metadata when the URL changes', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        hyperlink: 'https://old.example.test',
        hyperlinkDisplay: 'Old target',
        hyperlinkTooltip: 'Open old target',
      },
    );
    const handle = attachFormatDialog({ host, store });
    try {
      handle.open('more');
      const input = document.querySelector<HTMLInputElement>('input[aria-label="ハイパーリンク"]');
      if (!input) throw new Error('hyperlink input missing');
      input.value = 'https://new.example.test';
      input.dispatchEvent(new Event('input', { bubbles: true }));

      const ok = Array.from(
        document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__footer button'),
      ).find((b) => b.textContent === 'OK');
      ok?.click();

      const format = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
      expect(format?.hyperlink).toBe('https://new.example.test');
      expect(format?.hyperlinkDisplay).toBeUndefined();
      expect(format?.hyperlinkTooltip).toBeUndefined();
    } finally {
      handle.detach();
    }
  });

  it('preserves comment author metadata when applying Format Cells without clearing comments', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { comment: 'Keep this note', commentAuthor: 'Alice' },
    );
    const handle = attachFormatDialog({ host, store });
    try {
      handle.open('more');
      const ok = Array.from(
        document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__footer button'),
      ).find((b) => b.textContent === 'OK');
      ok?.click();

      const format = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
      expect(format).toMatchObject({ comment: 'Keep this note', commentAuthor: 'Alice' });
    } finally {
      handle.detach();
    }
  });

  it('clears comment author metadata when Format Cells clears the comment', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { comment: 'Delete this note', commentAuthor: 'Alice' },
    );
    const handle = attachFormatDialog({ host, store });
    try {
      handle.open('more');
      const comment = document.querySelector<HTMLTextAreaElement>(
        'textarea[aria-label="コメント"]',
      );
      if (!comment) throw new Error('comment textarea missing');
      comment.value = '';
      comment.dispatchEvent(new Event('input', { bubbles: true }));

      const ok = Array.from(
        document.querySelectorAll<HTMLButtonElement>('.fc-fmtdlg__footer button'),
      ).find((b) => b.textContent === 'OK');
      ok?.click();

      const format = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
      expect(format?.comment).toBeUndefined();
      expect(format?.commentAuthor).toBeUndefined();
    } finally {
      handle.detach();
    }
  });

  it('history records the format change as a single undo entry', () => {
    const history = new History();
    const handle = attachFormatDialog({ host, store, history });
    handle.open();

    const center = document.querySelector<HTMLInputElement>(
      'input[type="radio"][value="center"]',
    ) as HTMLInputElement;
    center.checked = true;
    center.dispatchEvent(new Event('change', { bubbles: true }));

    const bold = document.querySelector<HTMLInputElement>(
      'input[data-fc-check="bold"]',
    ) as HTMLInputElement;
    bold.checked = true;
    bold.dispatchEvent(new Event('change', { bubbles: true }));

    const okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    let fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.align).toBe('center');
    expect(fmt?.bold).toBe(true);

    expect(history.undo()).toBe(true);
    fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.align).toBeUndefined();
    expect(fmt?.bold).toBeFalsy();

    expect(history.redo()).toBe(true);
    fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.align).toBe('center');
    expect(fmt?.bold).toBe(true);

    handle.detach();
  });

  it('repeats a touched mixed value onto the current selection with F4 semantics', () => {
    const history = new History();
    setRange(store, 0, 0, 1, 0);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true, comment: 'source' });
    mutators.setCellFormat(store, { sheet: 0, row: 1, col: 0 }, { bold: false });
    const handle = attachFormatDialog({ host, store, history });
    handle.open('font');

    const bold = document.querySelector<HTMLInputElement>('input[data-fc-check="bold"]');
    if (!bold) throw new Error('bold checkbox missing');
    expect(bold.indeterminate).toBe(true);
    bold.checked = true;
    bold.dispatchEvent(new Event('change', { bubbles: true }));
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

    setRange(store, 0, 2, 1, 2);
    expect(history.repeatLast()).toBe(true);
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 2 }))?.bold).toBe(
      true,
    );
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 2 }))?.bold).toBe(
      true,
    );
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 2 }))?.comment,
    ).toBe(undefined);
    handle.detach();
  });

  it('keeps the repeat action after a format plus merge transaction', () => {
    const history = new History();
    setRange(store, 0, 0, 0, 1);
    const handle = attachFormatDialog({ host, store, history, getWb: mergeWorkbook });
    handle.open('align');
    const center = document.querySelector<HTMLInputElement>('input[type="radio"][value="center"]');
    const merge = document.querySelector<HTMLInputElement>('input[data-fc-check="mergeCells"]');
    if (!center || !merge) throw new Error('merge controls missing');
    center.checked = true;
    center.dispatchEvent(new Event('change', { bubbles: true }));
    merge.checked = true;
    merge.dispatchEvent(new Event('change', { bubbles: true }));
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

    expect(store.getState().merges.byAnchor.size).toBe(1);
    setRange(store, 0, 2, 0, 3);
    expect(history.repeatLast()).toBe(true);
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 2 }))?.align).toBe(
      'center',
    );
    expect(
      store.getState().merges.byAnchor.get(addrKey({ sheet: 0, row: 0, col: 2 })),
    ).toBeUndefined();
    handle.detach();
  });

  it('preserves the previous repeat when a same-value format merge aborts', () => {
    const history = new History();
    const previousRepeat = vi.fn();
    history.setRepeat(previousRepeat);
    setRange(store, 0, 0, 0, 1);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    setProtectedSheet(store, 0, true);
    const handle = attachFormatDialog({
      host,
      store,
      history,
      getWb: mergeWorkbook,
    });
    try {
      handle.open('font');
      const bold = document.querySelector<HTMLInputElement>('input[data-fc-check="bold"]');
      if (!bold) throw new Error('bold checkbox missing');
      expect(bold.checked).toBe(true);
      bold.dispatchEvent(new Event('change', { bubbles: true }));
      document
        .querySelector<HTMLButtonElement>('button[data-fc-tab="align"]')
        ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      const merge = document.querySelector<HTMLInputElement>('input[data-fc-check="mergeCells"]');
      if (!merge) throw new Error('merge checkbox missing');
      merge.checked = true;
      merge.dispatchEvent(new Event('change', { bubbles: true }));
      document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

      expect(history.repeatLast()).toBe(true);
      expect(previousRepeat).toHaveBeenCalledTimes(1);
      expect(store.getState().merges.byAnchor.size).toBe(0);
      expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.bold).toBe(
        true,
      );
    } finally {
      handle.detach();
    }
  });

  it('applies patch over the entire selection range', () => {
    setRange(store, 0, 0, 1, 1);
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const bold = document.querySelector<HTMLInputElement>(
      'input[data-fc-check="bold"]',
    ) as HTMLInputElement;
    bold.checked = true;
    bold.dispatchEvent(new Event('change', { bubbles: true }));

    const okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const formats = store.getState().format.formats;
    expect(formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.bold).toBe(true);
    expect(formats.get(addrKey({ sheet: 0, row: 0, col: 1 }))?.bold).toBe(true);
    expect(formats.get(addrKey({ sheet: 0, row: 1, col: 0 }))?.bold).toBe(true);
    expect(formats.get(addrKey({ sheet: 0, row: 1, col: 1 }))?.bold).toBe(true);
    handle.detach();
  });

  it('applies dialog format changes only to unlocked cells on protected sheets', () => {
    setRange(store, 0, 0, 0, 1);
    setCellLocked(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, false);
    setProtectedSheet(store, 0, true);
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const bold = document.querySelector<HTMLInputElement>(
      'input[data-fc-check="bold"]',
    ) as HTMLInputElement;
    bold.checked = true;
    bold.dispatchEvent(new Event('change', { bubbles: true }));

    const okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const formats = store.getState().format.formats;
    expect(formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.bold).toBeUndefined();
    expect(formats.get(addrKey({ sheet: 0, row: 0, col: 1 }))?.bold).toBe(true);
    handle.detach();
  });
});
