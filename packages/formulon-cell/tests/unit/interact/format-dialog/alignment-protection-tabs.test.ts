import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../../src/commands/interaction-controller.js';
import { fixedFormPolicy } from '../../../../src/commands/interaction-policy.js';
import { addrKey, WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { attachFormatDialog } from '../../../../src/interact/format-dialog.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';
import { mergeWorkbook, seedText, setActive, setRange } from './fixtures.js';

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

  it('follows the rotation with the dial marker and its sample text', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open('align');

    const dial = document.querySelector<HTMLElement>('.fc-fmtdlg__align-preview-dial');
    const pointer = document.querySelector<HTMLElement>('.fc-fmtdlg__align-preview-pointer');
    const sample = document.querySelector<HTMLElement>('.fc-fmtdlg__align-preview-text');
    const rotationInput = document.querySelector<HTMLInputElement>(
      '.fc-fmtdlg__align-degree input[type="number"]',
    );
    if (!dial || !pointer || !sample || !rotationInput) throw new Error('rotation dial missing');

    expect(sample.style.transform).toBe('translate(0, -50%) rotate(0deg)');

    dial.querySelector<HTMLButtonElement>('[data-fc-angle="60"]')?.click();
    expect(rotationInput.value).toBe('60');
    expect(sample.style.transform).toBe('translate(0, -50%) rotate(-60deg)');
    expect(pointer.style.top).not.toBe('66px');
    expect(
      dial.querySelector<HTMLElement>('.fc-fmtdlg__align-preview-dot--active')?.dataset.fcAngle,
    ).toBe('60');

    rotationInput.value = '-30';
    rotationInput.dispatchEvent(new Event('input', { bubbles: true }));
    expect(sample.style.transform).toBe('translate(0, -50%) rotate(30deg)');
    expect(
      dial.querySelector<HTMLElement>('.fc-fmtdlg__align-preview-dot--active')?.dataset.fcAngle,
    ).toBe('-30');

    handle.detach();
  });

  it('alignment radios update draft (default → undefined)', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const center = document.querySelector<HTMLInputElement>(
      'input[type="radio"][value="center"]',
    ) as HTMLInputElement;
    center.checked = true;
    center.dispatchEvent(new Event('change', { bubbles: true }));

    const okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.align).toBe(
      'center',
    );
    handle.detach();
  });

  it('Alignment tab exposes Excel-style extended horizontal and vertical choices', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    document
      .querySelector<HTMLButtonElement>('button[data-fc-tab="align"]')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const horizontal = document.querySelector<HTMLSelectElement>('select[data-fc-select="align"]');
    const vertical = document.querySelector<HTMLSelectElement>('select[data-fc-select="vAlign"]');
    if (!horizontal || !vertical) throw new Error('alignment selects missing');
    expect([...horizontal.options].map((o) => o.value)).toEqual([
      'default',
      'left',
      'center',
      'right',
      'fill',
      'justify',
      'centerContinuous',
      'distributed',
    ]);
    expect([...vertical.options].map((o) => o.value)).toEqual([
      'default',
      'top',
      'middle',
      'bottom',
      'justify',
      'distributed',
    ]);

    const justifyLastLine = document.querySelector<HTMLInputElement>(
      'input[data-fc-check="justifyLastLine"]',
    );
    expect(justifyLastLine?.disabled).toBe(true);
    horizontal.value = 'distributed';
    horizontal.dispatchEvent(new Event('change', { bubbles: true }));
    expect(justifyLastLine?.disabled).toBe(false);

    horizontal.value = 'centerContinuous';
    horizontal.dispatchEvent(new Event('change', { bubbles: true }));
    vertical.value = 'distributed';
    vertical.dispatchEvent(new Event('change', { bubbles: true }));

    document
      .querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.align).toBe('centerContinuous');
    expect(fmt?.vAlign).toBe('distributed');
    handle.detach();
  });

  it('Alignment tab persists Shrink to Fit like Excel text control', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    document
      .querySelector<HTMLButtonElement>('button[data-fc-tab="align"]')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const shrink = document.querySelector<HTMLInputElement>('input[data-fc-check="shrinkToFit"]');
    expect(shrink?.disabled).toBe(false);
    if (!shrink) throw new Error('shrink checkbox missing');
    shrink.checked = true;
    shrink.dispatchEvent(new Event('change', { bubbles: true }));

    document
      .querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.shrinkToFit,
    ).toBe(true);
    handle.detach();
  });

  it('Alignment tab persists text direction', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    document
      .querySelector<HTMLButtonElement>('button[data-fc-tab="align"]')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const direction = document.querySelector<HTMLSelectElement>(
      'select[data-fc-select="textDirection"]',
    );
    if (!direction) throw new Error('text direction select missing');
    direction.value = 'rtl';
    direction.dispatchEvent(new Event('change', { bubbles: true }));

    document
      .querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.textDirection,
    ).toBe('rtl');
    handle.detach();
  });

  it('Alignment tab Merge cells applies a merge to a multi-cell selection', () => {
    setRange(store, 0, 0, 1, 1);
    const handle = attachFormatDialog({ host, store });
    handle.open();
    document
      .querySelector<HTMLButtonElement>('button[data-fc-tab="align"]')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const merge = document.querySelector<HTMLInputElement>('input[data-fc-check="mergeCells"]');
    expect(merge?.disabled).toBe(false);
    if (!merge) throw new Error('merge checkbox missing');
    merge.checked = true;
    merge.dispatchEvent(new Event('change', { bubbles: true }));

    document
      .querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(store.getState().merges.byAnchor.get(addrKey({ sheet: 0, row: 0, col: 0 }))).toEqual({
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 1,
      c1: 1,
    });
    handle.detach();
  });

  it('rejects a restricted format plus merge action before either mutation', async () => {
    setRange(store, 0, 0, 0, 1);
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const history = new History();
    const controller = new InteractionController({
      store,
      getWb: () => workbook,
      history,
    });
    const unregister = registerInteractionController(store, controller);
    controller.setPolicy({
      ...fixedFormPolicy([{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 }]),
      operations: { format: true, merge: true },
    });
    const handle = attachFormatDialog({ host, store, history, getWb: () => workbook });
    try {
      handle.open('align');
      const center = document.querySelector<HTMLInputElement>(
        'input[type="radio"][value="center"]',
      );
      const merge = document.querySelector<HTMLInputElement>('input[data-fc-check="mergeCells"]');
      if (!center || !merge) throw new Error('merge controls missing');
      center.checked = true;
      center.dispatchEvent(new Event('change', { bubbles: true }));
      merge.checked = true;
      merge.dispatchEvent(new Event('change', { bubbles: true }));
      document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

      expect(store.getState().format.formats.size).toBe(0);
      expect(store.getState().merges.byAnchor.size).toBe(0);
      expect(document.querySelector<HTMLElement>('.fc-fmtdlg')?.hidden).toBe(false);
    } finally {
      handle.detach();
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  it('Alignment tab Merge cells projects a disabled reason for a single unmerged cell', () => {
    setActive(store, 0, 0);
    const handle = attachFormatDialog({ host, store });
    handle.open();
    document
      .querySelector<HTMLButtonElement>('button[data-fc-tab="align"]')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const merge = document.querySelector<HTMLInputElement>('input[data-fc-check="mergeCells"]');
    if (!merge) throw new Error('merge checkbox missing');
    const wrap = merge.closest<HTMLElement>('label');
    expect(merge.disabled).toBe(true);
    expect(merge.dataset.disabledReason).toBe(
      'セルを結合するには 2 つ以上のセルを選択してください。',
    );
    expect(wrap?.dataset.disabledReason).toBe(
      'セルを結合するには 2 つ以上のセルを選択してください。',
    );
    handle.detach();
  });

  it('Alignment tab Merge cells hydrates checked and can unmerge', () => {
    setRange(store, 0, 0, 1, 1);
    mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    const handle = attachFormatDialog({ host, store });
    handle.open();
    document
      .querySelector<HTMLButtonElement>('button[data-fc-tab="align"]')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const merge = document.querySelector<HTMLInputElement>('input[data-fc-check="mergeCells"]');
    expect(merge?.checked).toBe(true);
    if (!merge) throw new Error('merge checkbox missing');
    merge.checked = false;
    merge.dispatchEvent(new Event('change', { bubbles: true }));

    document
      .querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(store.getState().merges.byAnchor.size).toBe(0);
    handle.detach();
  });

  it('warns before format or merge mutations and keeps the dialog open on cancel', () => {
    setRange(store, 0, 0, 0, 1);
    seedText(store, 0, 0, 'anchor');
    seedText(store, 0, 1, 'other');
    const handle = attachFormatDialog({ host, store, getWb: () => mergeWorkbook() });
    handle.open('align');

    const merge = document.querySelector<HTMLInputElement>('input[data-fc-check="mergeCells"]');
    if (!merge) throw new Error('merge checkbox missing');
    merge.checked = true;
    merge.dispatchEvent(new Event('change', { bubbles: true }));
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

    expect(document.querySelector<HTMLElement>('.fc-fmtdlg')?.hidden).toBe(false);
    expect(store.getState().merges.byAnchor.size).toBe(0);
    expect(store.getState().format.formats.size).toBe(0);

    const cancel = [
      ...document.querySelectorAll<HTMLButtonElement>('[role="alertdialog"] button'),
    ].find((button) => button.textContent === 'キャンセル');
    cancel?.click();
    expect(document.querySelector<HTMLElement>('.fc-fmtdlg')?.hidden).toBe(false);
    expect(store.getState().merges.byAnchor.size).toBe(0);
    expect(store.getState().format.formats.size).toBe(0);

    handle.detach();
  });

  it('leaves mixed merges untouched when the checkbox is unchanged', () => {
    setRange(store, 0, 0, 0, 3);
    mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 3 });
    const handle = attachFormatDialog({ host, store });
    handle.open('align');

    const merge = document.querySelector<HTMLInputElement>('input[data-fc-check="mergeCells"]');
    if (!merge) throw new Error('merge checkbox missing');
    expect(merge.checked).toBe(false);
    expect(merge.indeterminate).toBe(true);
    document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

    expect(store.getState().merges.byAnchor.size).toBe(2);
    handle.detach();
  });

  it('Protection tab persists the Hidden formula flag', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    document
      .querySelector<HTMLButtonElement>('button[data-fc-tab="protection"]')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const hidden = document.querySelector<HTMLInputElement>('input[data-fc-check="formulaHidden"]');
    if (!hidden) throw new Error('formula hidden checkbox missing');
    const panel = document.querySelector<HTMLDivElement>('div[data-fc-tab="protection"]');
    const hint = panel?.querySelector<HTMLDivElement>('.fc-fmtdlg__protection-hint');
    const checks = panel?.querySelector<HTMLDivElement>('.fc-fmtdlg__protection-checks');
    expect(hint?.textContent).toContain('ワークシートを保護します');
    expect(checks?.children.length).toBe(2);
    expect(hint && checks ? hint.compareDocumentPosition(checks) : 0).toBe(
      Node.DOCUMENT_POSITION_FOLLOWING,
    );
    hidden.checked = true;
    hidden.dispatchEvent(new Event('change', { bubbles: true }));

    document
      .querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.formulaHidden,
    ).toBe(true);
    handle.detach();
  });

  it('alignment "default" radio clears align', () => {
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { align: 'right' });
    const handle = attachFormatDialog({ host, store });
    handle.open();

    const dflt = document.querySelector<HTMLInputElement>(
      'input[type="radio"][value="default"]',
    ) as HTMLInputElement;
    dflt.checked = true;
    dflt.dispatchEvent(new Event('change', { bubbles: true }));

    const okBtn = document.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary');
    okBtn?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }))?.align,
    ).toBeUndefined();
    handle.detach();
  });

  it('unchecked alignment radio change is ignored', () => {
    const handle = attachFormatDialog({ host, store });
    handle.open();
    const left = document.querySelector<HTMLInputElement>(
      'input[type="radio"][value="left"]',
    ) as HTMLInputElement;
    left.checked = false;
    left.dispatchEvent(new Event('change', { bubbles: true }));
    handle.detach();
  });
});
