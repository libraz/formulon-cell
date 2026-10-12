import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { setFont } from '../../src/commands/format.js';
import { moveSheet, removeSheet } from '../../src/commands/sheet-mutate.js';
import { MAX_COL, MAX_ROW } from '../../src/engine/address.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { createDefaultRibbonHelpers } from '../../src/mount/toolbar-defaults.js';
import { Spreadsheet } from '../../src/mount.js';
import { mutators } from '../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/mount.js';

const a1 = { sheet: 0, row: 0, col: 0 };
const far = { sheet: 0, row: MAX_ROW, col: MAX_COL };

const fontAt = (wb: WorkbookHandle, addr = a1) => {
  const style = wb.getEffectiveStyle(addr.sheet, addr.row, addr.col);
  const xf = style && wb.getCellXf(style.xfIndex);
  const font = xf && wb.getFontRecord(xf.fontIndex);
  expect(font).not.toBeNull();
  return font;
};

describe('real WASM whole-sheet fonts', () => {
  let mounted: MountedStubSheet;
  let toolbar: ReturnType<typeof Spreadsheet.mountToolbar>;
  let toolbarHost: HTMLElement;

  beforeEach(async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'ja' });
    expect(workbook.isStub).toBe(false);
    mounted = await mountStubSheet({ workbook, locale: 'ja' });
    toolbarHost = document.createElement('div');
    document.body.append(toolbarHost);
    toolbar = Spreadsheet.mountToolbar(toolbarHost, mounted.instance, {
      helpers: createDefaultRibbonHelpers(mounted.instance, { lang: 'ja' }),
    });
  });

  afterEach(() => {
    toolbar.dispose();
    mounted.dispose();
    mounted.workbook.dispose();
    toolbarHost.remove();
  });

  const chooseSize = (points: number): void => {
    toolbarHost
      .querySelector<HTMLButtonElement>('[data-ribbon-command="fontSize"] button')
      ?.click();
    const option = document.querySelector<HTMLButtonElement>(
      `.fc-tb__rb-dd__opt[data-value="${points}"]`,
    );
    expect(option).not.toBeNull();
    option?.click();
  };

  it('changes empty distant cells sparsely and undoes/redoes an empty worksheet', () => {
    const { instance, workbook } = mounted;
    mutators.selectAll(instance.store);
    chooseSize(18);
    expect(fontAt(workbook)?.size).toBe(18);
    expect(fontAt(workbook, far)?.size).toBe(18);
    expect(instance.store.getState().format.formats.size).toBe(0);
    expect(instance.store.getState().data.cells.size).toBe(0);
    expect(instance.undo()).toBe(true);
    expect(fontAt(workbook, far)?.size).toBe(11);
    expect(instance.redo()).toBe(true);
    expect(fontAt(workbook, far)?.size).toBe(18);
  });

  it('keeps individual overrides and future input through save/reload', async () => {
    const { instance, workbook } = mounted;
    expect(instance.applyChanges([{ addr: a1, input: 'existing' }]).status).toBe('applied');
    mutators.selectAll(instance.store);
    chooseSize(18);
    mutators.setActive(instance.store, a1);
    chooseSize(24);
    expect(instance.applyChanges([{ addr: { ...a1, col: 1 }, input: 'new' }]).status).toBe(
      'applied',
    );
    expect(fontAt(workbook)?.size).toBe(24);
    expect(fontAt(workbook, { ...a1, col: 1 })?.size).toBe(18);
    expect(fontAt(workbook, far)?.size).toBe(18);
    const reloaded = await WorkbookHandle.loadBytes(workbook.save());
    const reloadedMount = await mountStubSheet({ workbook: reloaded, locale: 'ja' });
    try {
      expect(fontAt(reloaded)?.size).toBe(24);
      expect(fontAt(reloaded, { ...a1, col: 1 })?.size).toBe(18);
      expect(fontAt(reloaded, far)?.size).toBe(18);
      expect(
        reloadedMount.instance.applyChanges([{ addr: { ...a1, col: 2 }, input: 'after reload' }])
          .status,
      ).toBe('applied');
      expect(fontAt(reloaded, { ...a1, col: 2 })?.size).toBe(18);
    } finally {
      reloadedMount.dispose();
      reloaded.dispose();
    }
  });

  it('keeps other sheets unchanged and follows sheet move/remove', async () => {
    const { instance, workbook } = mounted;
    workbook.addSheet('Other');
    mutators.selectAll(instance.store);
    chooseSize(18);
    expect(fontAt(workbook, { ...far, sheet: 1 })?.size).toBe(11);
    expect(moveSheet(instance.store, workbook, 0, 1)).toBe(true);
    expect(fontAt(workbook, { ...far, sheet: 1 })?.size).toBe(18);
    expect(fontAt(workbook, far)?.size).toBe(11);
    expect(removeSheet(instance.store, workbook, 0)).toBe(true);
    expect(fontAt(workbook, far)?.size).toBe(18);
    const reloaded = await WorkbookHandle.loadBytes(workbook.save());
    try {
      expect(reloaded.sheetCount).toBe(1);
      expect(fontAt(reloaded, far)?.size).toBe(18);
    } finally {
      reloaded.dispose();
    }
  });

  it('retains inherited raw font metadata when entering a new cell after reload', async () => {
    const wb = mounted.workbook;
    const baseFont = wb.getFontRecord(0);
    const baseXf = wb.getCellXf(0);
    expect(baseFont).not.toBeNull();
    expect(baseXf).not.toBeNull();
    if (!baseFont || !baseXf) throw new Error('Missing default style');
    const sourceFont = {
      ...baseFont,
      size: 18,
      bold: true,
      hasBold: true,
      charset: 128,
      hasCharset: true,
      color: { kind: 2, rgb: 0, theme: 4, tint: 0.4, indexed: 0 },
    };
    const xfIndex = wb.addXfRecord({ ...baseXf, fontIndex: wb.addFontRecord(sourceFont) });
    expect(wb.setColumnStyle(0, 0, MAX_COL, xfIndex)).toBe(true);
    const reloaded = await WorkbookHandle.loadBytes(wb.save());
    const reloadedMount = await mountStubSheet({ workbook: reloaded, locale: 'ja' });
    try {
      expect(fontAt(reloaded, far)).toMatchObject(sourceFont);
      expect(reloadedMount.instance.applyChanges([{ addr: a1, input: 'axis style' }]).status).toBe(
        'applied',
      );
      expect(fontAt(reloaded)).toMatchObject(sourceFont);
      expect(
        setFont(reloadedMount.instance.store.getState(), reloadedMount.instance.store, {
          fontSize: 32,
        }),
      ).toBe(true);
      expect(fontAt(reloaded)).toMatchObject({ ...sourceFont, size: 24 });
    } finally {
      reloadedMount.dispose();
      reloaded.dispose();
    }
  });

  it('applies a font-only full-sheet Format Cells change and supports undo', () => {
    const { instance, workbook } = mounted;
    mutators.selectAll(instance.store);
    instance.openFormatDialog('font');
    const input = document.querySelector<HTMLInputElement>(
      '.fc-fmtdlg:not([hidden]) [data-fc-input="size"]',
    );
    expect(input).not.toBeNull();
    if (!input) throw new Error('Missing font size input');
    input.value = '18';
    input.dispatchEvent(new Event('input', { bubbles: true }));
    const ok = document.querySelector<HTMLButtonElement>(
      '.fc-fmtdlg:not([hidden]) .fc-fmtdlg__footer .fc-fmtdlg__btn--primary',
    );
    expect(ok).not.toBeNull();
    ok?.click();
    expect(instance.store.getState().format.sheetFontOverrides?.get(0)?.patch.fontSize).toBe(24);
    expect(fontAt(workbook, far)?.size).toBe(18);
    expect(instance.undo()).toBe(true);
    expect(fontAt(workbook, far)?.size).toBe(11);
  });

  it('rejects the entire font change on a protected worksheet', () => {
    const { instance, workbook } = mounted;
    instance.store.setState((state) => ({
      ...state,
      protection: { ...state.protection, protectedSheets: new Map([[0, {}]]) },
    }));
    mutators.selectAll(instance.store);
    expect(setFont(instance.store.getState(), instance.store, { fontSize: 24 })).toBe(false);
    expect(instance.store.getState().format.sheetFontOverrides?.size ?? 0).toBe(0);
    expect(fontAt(workbook, far)?.size).toBe(11);
  });
});
