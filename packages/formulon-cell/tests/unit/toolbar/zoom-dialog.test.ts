import { afterEach, describe, expect, it } from 'vitest';

import { SheetKind } from '../../../src/engine/types.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import {
  selectionZoom,
  showMacZoomDialog,
  showZoomDialog,
  zoomForSelection,
} from '../../../src/toolbar/dialogs/zoom.js';

describe('showZoomDialog', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('returns a bounded zoom percentage', async () => {
    const pending = showZoomDialog({
      title: 'Zoom',
      label: 'Magnification',
      initial: 100,
      okLabel: 'OK',
      cancelLabel: 'Cancel',
      invalidMessage: 'Enter a zoom percentage from 10 to 400.',
    });

    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(dialog?.textContent).toContain('Magnification');
    const input = dialog?.querySelector<HTMLInputElement>('input[type="number"]');
    expect(input?.min).toBe('10');
    expect(input?.max).toBe('400');
    expect(input?.step).toBe('1');
    if (!input) throw new Error('Expected zoom input.');
    input.value = '125';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

    await expect(pending).resolves.toBe(125);
  });

  it('shows the supplied validation message instead of reusing the field label', async () => {
    const pending = showZoomDialog({
      title: 'Zoom',
      label: 'Magnification',
      initial: 100,
      okLabel: 'OK',
      cancelLabel: 'Cancel',
      invalidMessage: 'Enter a zoom percentage from 10 to 400.',
    });

    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    const input = dialog?.querySelector<HTMLInputElement>('input[type="number"]');
    if (!input) throw new Error('Expected zoom input.');
    input.value = '401';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

    expect(dialog?.textContent).toContain('Enter a zoom percentage from 10 to 400.');
    input.value = '150';
    dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();
    await expect(pending).resolves.toBe(150);
  });
});

describe('showMacZoomDialog', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  const open = (initial: number) =>
    showMacZoomDialog({
      title: 'ズーム',
      caption: '倍率:',
      fitSelection: '選択範囲をズーム',
      custom: 'カスタム:',
      initial,
      okLabel: 'OK',
      cancelLabel: 'キャンセル',
      invalidMessage: '10 から 400 までの倍率を入力してください。',
    });
  const dialog = (): HTMLElement => {
    const el = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    if (!el) throw new Error('missing zoom dialog');
    return el;
  };
  const radios = (): HTMLInputElement[] => [
    ...dialog().querySelectorAll<HTMLInputElement>('input[type="radio"]'),
  ];
  const ok = (): void =>
    dialog().querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')?.click();

  it('offers the desktop presets, zoom to selection and a custom percentage', async () => {
    const pending = open(100);
    expect(radios().map((r) => r.parentElement?.textContent)).toEqual([
      '200%',
      '150%',
      '125%',
      '100%',
      '75%',
      '50%',
      '25%',
      '選択範囲をズーム',
      'カスタム:',
    ]);
    expect(dialog().textContent).toContain('倍率:');
    expect(radios()[3]?.checked).toBe(true);
    const custom = dialog().querySelector<HTMLInputElement>('input[type="number"]');
    expect(custom?.value).toBe('100');
    radios()[1]?.click();
    ok();
    await expect(pending).resolves.toEqual({ kind: 'percent', percent: 150 });
  });

  it('returns zoom to selection', async () => {
    const pending = open(100);
    radios()[7]?.click();
    ok();
    await expect(pending).resolves.toEqual({ kind: 'selection' });
  });

  it('starts on custom for a non-preset zoom and validates the field', async () => {
    const pending = open(110);
    expect(radios()[8]?.checked).toBe(true);
    const custom = dialog().querySelector<HTMLInputElement>('input[type="number"]');
    if (!custom) throw new Error('missing custom field');
    custom.value = '401';
    ok();
    expect(dialog().textContent).toContain('10 から 400 までの倍率を入力してください。');
    custom.value = '333';
    ok();
    await expect(pending).resolves.toEqual({ kind: 'percent', percent: 333 });
  });
});

describe('zoomForSelection', () => {
  it('fits 20 rows or 12 columns and clamps to the zoom range', () => {
    expect(zoomForSelection({ sheet: 0, r0: 0, c0: 0, r1: 39, c1: 0 })).toBe(0.5);
    expect(zoomForSelection({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 23 })).toBe(0.5);
    expect(zoomForSelection({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 })).toBe(4);
  });
});

describe('selectionZoom', () => {
  it('fits the selection on a worksheet', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    expect(selectionZoom(store.getState())).toBe(4);
  });

  it('takes the minimum zoom on a chartsheet, which has no cells to fit', () => {
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, sheetKinds: new Map([[0, SheetKind.Chartsheet]]) },
    }));
    expect(selectionZoom(store.getState())).toBe(0.1);
  });
});
