import { describe, expect, it } from 'vitest';
import { applyCellStyleToSelection } from '../../src/commands/cell-styles.js';
import { recordDialogFormatChange } from '../../src/commands/dialog-format-history.js';
import { applySelectionFormatAction, setFont, setFontColor } from '../../src/commands/format.js';
import { History } from '../../src/commands/history.js';
import { recordFormatChange } from '../../src/commands/slice-history.js';
import { MAX_COL, MAX_ROW } from '../../src/engine/address.js';
import { patchXfIndexFromBase } from '../../src/engine/cell-xf-writeback.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { createDefaultRibbonHelpers } from '../../src/mount/toolbar-defaults.js';
import { Spreadsheet } from '../../src/mount.js';
import { mutators } from '../../src/store/store.js';
import { mountStubSheet } from '../test-utils/mount.js';

describe('real WASM axis selection font size', () => {
  for (const axis of ['column', 'rows'] as const) {
    it(`changes ${axis} including distant blanks and preserves adjacent cells`, async () => {
      const workbook = await WorkbookHandle.createDefault({ locale: 'ja' });
      const mounted = await mountStubSheet({ workbook, locale: 'ja' });
      const host = document.createElement('div');
      document.body.append(host);
      const toolbar = Spreadsheet.mountToolbar(host, mounted.instance, {
        helpers: createDefaultRibbonHelpers(mounted.instance, { lang: 'ja' }),
      });
      const sizeAt = (row: number, col: number): number | undefined => {
        const effective = workbook.getEffectiveStyle(0, row, col);
        const xf = effective && workbook.getCellXf(effective.xfIndex);
        return xf ? workbook.getFontRecord(xf.fontIndex)?.size : undefined;
      };
      try {
        expect(workbook.isStub).toBe(false);
        if (axis === 'column') mutators.selectCol(mounted.instance.store, 1);
        else mutators.selectRows(mounted.instance.store, 2, 9);
        host.querySelector<HTMLButtonElement>('[data-ribbon-command="fontSize"] button')?.click();
        const option = document.querySelector<HTMLButtonElement>(
          '.fc-tb__rb-dd__opt[data-value="18"]',
        );
        expect(option).not.toBeNull();
        option?.click();
        const target = axis === 'column' ? [MAX_ROW, 1] : [9, MAX_COL];
        expect(sizeAt(target[0] as number, target[1] as number)).toBe(18);
        expect(sizeAt(0, 0)).toBe(11);
        expect(mounted.instance.store.getState().format.formats.size).toBeLessThan(100);
        expect(mounted.instance.undo()).toBe(true);
        expect(sizeAt(target[0] as number, target[1] as number)).toBe(11);
        expect(mounted.instance.redo()).toBe(true);
        expect(sizeAt(target[0] as number, target[1] as number)).toBe(18);
      } finally {
        toolbar.dispose();
        mounted.dispose();
        workbook.dispose();
        host.remove();
      }
    });
  }

  it('preserves a varying column font when a row font size is applied', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'ja' });
    const mounted = await mountStubSheet({ workbook, locale: 'ja' });
    const fontAt = (row: number, col: number) => {
      const style = workbook.getEffectiveStyle(0, row, col);
      const xf = style && workbook.getCellXf(style.xfIndex);
      return xf ? workbook.getFontRecord(xf.fontIndex) : null;
    };
    try {
      expect(workbook.isStub).toBe(false);
      mutators.selectCol(mounted.instance.store, 1);
      setFontColor(mounted.instance.store.getState(), mounted.instance.store, '#ff0000');
      const columnFont = fontAt(3, 1);
      expect(columnFont?.colorArgb).toBe(0xffff0000);

      mutators.selectRows(mounted.instance.store, 3, 3);
      expect(
        setFont(mounted.instance.store.getState(), mounted.instance.store, { fontSize: 24 }),
      ).toBe(true);
      expect(fontAt(3, 0)?.size).toBe(18);
      expect(fontAt(3, 1)).toMatchObject({ size: 18, colorArgb: columnFont?.colorArgb });
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('uses the widest column style as the row base when column A is the exception', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'ja' });
    const mounted = await mountStubSheet({ workbook, locale: 'ja' });
    const fontAt = (row: number, col: number) => {
      const style = workbook.getEffectiveStyle(0, row, col);
      const xf = style && workbook.getCellXf(style.xfIndex);
      return xf ? workbook.getFontRecord(xf.fontIndex) : null;
    };
    try {
      expect(workbook.isStub).toBe(false);
      mutators.selectCol(mounted.instance.store, 0);
      setFontColor(mounted.instance.store.getState(), mounted.instance.store, '#ff0000');
      const columnFont = fontAt(3, 0);
      expect(columnFont?.colorArgb).toBe(0xffff0000);

      mutators.selectRows(mounted.instance.store, 3, 10);
      expect(
        setFont(mounted.instance.store.getState(), mounted.instance.store, { fontSize: 24 }),
      ).toBe(true);
      expect(fontAt(3, 0)).toMatchObject({ size: 18, colorArgb: columnFont?.colorArgb });
      expect(fontAt(10, MAX_COL)?.size).toBe(18);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('keeps the default gap when the widest column segment is styled', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'ja' });
    const mounted = await mountStubSheet({ workbook, locale: 'ja' });
    const fontAt = (row: number, col: number) => {
      const style = workbook.getEffectiveStyle(0, row, col);
      const xf = style && workbook.getCellXf(style.xfIndex);
      return xf ? workbook.getFontRecord(xf.fontIndex) : null;
    };
    try {
      expect(workbook.isStub).toBe(false);
      mutators.selectCols(mounted.instance.store, 1, MAX_COL);
      setFontColor(mounted.instance.store.getState(), mounted.instance.store, '#ff0000');
      expect(fontAt(3, 1)?.colorArgb).toBe(0xffff0000);

      mutators.selectRow(mounted.instance.store, 3);
      expect(
        setFont(mounted.instance.store.getState(), mounted.instance.store, { fontSize: 24 }),
      ).toBe(true);
      expect(fontAt(3, 1)).toMatchObject({ size: 18, colorArgb: 0xffff0000 });
      expect(fontAt(3, 0)).toMatchObject({ size: 18, colorArgb: 0xff000000 });
      expect(fontAt(2, 1)).toMatchObject({ size: 11, colorArgb: 0xffff0000 });
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('keeps the full-sheet and column override history sequence reversible', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'ja' });
    const mounted = await mountStubSheet({ workbook, locale: 'ja' });
    const history = new History();
    const sizeAt = (row: number, col: number): number | undefined => {
      const effective = workbook.getEffectiveStyle(0, row, col);
      const xf = effective && workbook.getCellXf(effective.xfIndex);
      return xf ? workbook.getFontRecord(xf.fontIndex)?.size : undefined;
    };
    const setSize = (size: number): void => {
      recordFormatChange(history, mounted.instance.store, () => {
        setFont(mounted.instance.store.getState(), mounted.instance.store, { fontSize: size });
      });
    };
    try {
      mutators.selectCol(mounted.instance.store, 1);
      setSize(24);
      expect(sizeAt(0, 1)).toBe(18);

      mutators.selectAll(mounted.instance.store);
      setSize(32);
      expect(sizeAt(0, 0)).toBe(24);
      expect(sizeAt(0, 1)).toBe(24);

      mutators.selectCol(mounted.instance.store, 1);
      setSize(40);
      expect(sizeAt(0, 1)).toBe(30);

      expect(history.undo()).toBe(true);
      expect(sizeAt(0, 0)).toBe(24);
      expect(sizeAt(0, 1)).toBe(24);
      expect(history.undo()).toBe(true);
      expect(sizeAt(0, 0)).toBe(11);
      expect(sizeAt(0, 1)).toBe(18);
      expect(history.undo()).toBe(true);
      expect(sizeAt(0, 0)).toBe(11);
      expect(sizeAt(0, 1)).toBe(11);

      expect(history.redo()).toBe(true);
      expect(sizeAt(0, 1)).toBe(18);
      expect(history.redo()).toBe(true);
      expect(sizeAt(0, 0)).toBe(24);
      expect(sizeAt(0, 1)).toBe(24);
      expect(history.redo()).toBe(true);
      expect(sizeAt(0, 1)).toBe(30);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('keeps an excluded hole unchanged when a column union uses native axis style', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'ja' });
    const mounted = await mountStubSheet({ workbook, locale: 'ja' });
    const history = new History();
    const sizeAt = (row: number, col: number): number | undefined => {
      const effective = workbook.getEffectiveStyle(0, row, col);
      const xf = effective && workbook.getCellXf(effective.xfIndex);
      return xf ? workbook.getFontRecord(xf.fontIndex)?.size : undefined;
    };
    try {
      mounted.instance.store.setState((state) => ({
        ...state,
        selection: {
          ...state.selection,
          range: { sheet: 0, r0: 0, r1: 0, c0: 1, c1: 1 },
          extraRanges: [{ sheet: 0, r0: 2, r1: MAX_ROW, c0: 1, c1: 1 }],
        },
      }));
      recordFormatChange(history, mounted.instance.store, () => {
        if (
          !applySelectionFormatAction(mounted.instance.store.getState(), mounted.instance.store, {
            patch: { fontSize: 24 },
          })
        )
          throw new Error('axis hole format was rejected');
      });
      expect(sizeAt(0, 1)).toBe(18);
      expect(sizeAt(1, 1)).toBe(11);
      expect(sizeAt(MAX_ROW, 1)).toBe(18);
      expect(mounted.instance.store.getState().format.formats.has('0:1:1')).toBe(false);
      expect(history.undo()).toBe(true);
      expect(sizeAt(0, 1)).toBe(11);
      expect(sizeAt(1, 1)).toBe(11);
      expect(sizeAt(MAX_ROW, 1)).toBe(11);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('does not reflatten a derived cell cleared back to its row source', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'ja' });
    const mounted = await mountStubSheet({ workbook, locale: 'ja' });
    const appearance = () => {
      const effective = workbook.getEffectiveStyle(0, 2, 1);
      const xf = effective && workbook.getCellXf(effective.xfIndex);
      const font = xf && workbook.getFontRecord(xf.fontIndex);
      return { source: effective?.source, size: font?.size, colorArgb: font?.colorArgb };
    };
    try {
      const redXf = patchXfIndexFromBase(workbook, 0, { color: '#ff0000' });
      expect(redXf).toBeGreaterThan(0);
      expect(workbook.setRowStyle(0, 2, redXf)).toBe(true);

      mutators.selectCol(mounted.instance.store, 1);
      expect(
        setFont(mounted.instance.store.getState(), mounted.instance.store, { fontSize: 24 }),
      ).toBe(true);
      expect(appearance()).toMatchObject({ source: 0, size: 18, colorArgb: 0xffff0000 });

      // Clearing the derived direct XF restores row inheritance for a blank
      // cell. A later unrelated axis publication must preserve that source.
      expect(workbook.clearCellXfIndex(0, 2, 1)).toBe(true);
      expect(appearance()).toMatchObject({ source: 1, size: 11, colorArgb: 0xffff0000 });

      mutators.selectCol(mounted.instance.store, 2);
      expect(
        applySelectionFormatAction(mounted.instance.store.getState(), mounted.instance.store, {
          patch: { bold: true },
        }),
      ).toBe(true);
      expect(appearance()).toMatchObject({ source: 1, size: 11, colorArgb: 0xffff0000 });
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('records a dialog axis edit without expanding its target cells', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'ja' });
    const mounted = await mountStubSheet({ workbook, locale: 'ja' });
    const history = new History();
    const sizeAt = (row: number, col: number): number | undefined => {
      const effective = workbook.getEffectiveStyle(0, row, col);
      const xf = effective && workbook.getCellXf(effective.xfIndex);
      return xf ? workbook.getFontRecord(xf.fontIndex)?.size : undefined;
    };
    try {
      mutators.selectCol(mounted.instance.store, 1);
      expect(
        recordDialogFormatChange({
          history,
          store: mounted.instance.store,
          workbook,
          sheet: 0,
          targets: [],
          pendingBefore: null,
          mutate: () =>
            applySelectionFormatAction(mounted.instance.store.getState(), mounted.instance.store, {
              patch: { fontSize: 24 },
            }),
        }),
      ).toBe(true);
      expect(sizeAt(MAX_ROW, 1)).toBe(18);
      expect(history.undo()).toBe(true);
      expect(sizeAt(MAX_ROW, 1)).toBe(11);
      expect(history.redo()).toBe(true);
      expect(sizeAt(MAX_ROW, 1)).toBe(18);
      expect(mounted.instance.store.getState().format.formats.size).toBeLessThan(100);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('restores the legacy whole-sheet font around a dialog axis edit', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'ja' });
    const mounted = await mountStubSheet({ workbook, locale: 'ja' });
    const history = new History();
    try {
      mutators.selectAll(mounted.instance.store);
      recordFormatChange(history, mounted.instance.store, () => {
        if (!setFont(mounted.instance.store.getState(), mounted.instance.store, { fontSize: 24 }))
          throw new Error('whole-sheet font format was rejected');
      });
      expect(mounted.instance.store.getState().format.sheetFontOverrides?.has(0)).toBe(true);

      mutators.selectCol(mounted.instance.store, 1);
      expect(
        recordDialogFormatChange({
          history,
          store: mounted.instance.store,
          workbook,
          sheet: 0,
          targets: [],
          pendingBefore: null,
          mutate: () =>
            applySelectionFormatAction(mounted.instance.store.getState(), mounted.instance.store, {
              patch: { fontSize: 24 },
            }),
        }),
      ).toBe(true);
      expect(mounted.instance.store.getState().format.sheetFontOverrides?.has(0)).toBe(false);
      expect(mounted.instance.store.getState().format.sheetAxisOverrides?.has(0)).toBe(true);

      expect(history.undo()).toBe(true);
      expect(mounted.instance.store.getState().format.sheetFontOverrides?.has(0)).toBe(true);
      expect(mounted.instance.store.getState().format.sheetAxisOverrides?.has(0)).toBe(false);
      expect(history.undo()).toBe(true);
      expect(mounted.instance.store.getState().format.sheetFontOverrides?.has(0)).toBe(false);

      expect(history.redo()).toBe(true);
      expect(mounted.instance.store.getState().format.sheetFontOverrides?.has(0)).toBe(true);
      expect(history.redo()).toBe(true);
      expect(mounted.instance.store.getState().format.sheetFontOverrides?.has(0)).toBe(false);
      expect(mounted.instance.store.getState().format.sheetAxisOverrides?.has(0)).toBe(true);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('routes a cell style gallery action through the compact axis plan', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'ja' });
    const mounted = await mountStubSheet({ workbook, locale: 'ja' });
    try {
      mutators.selectCol(mounted.instance.store, 1);
      const styleResult = applyCellStyleToSelection(
        mounted.instance.store,
        mounted.instance.history,
        'good',
        { getWorkbook: () => workbook },
      );
      expect(styleResult).toBe(true);
      expect(mounted.instance.store.getState().format.formats.size).toBeLessThan(100);
      expect(mounted.instance.store.getState().format.sheetAxisOverrides?.has(0)).toBe(true);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });
});
