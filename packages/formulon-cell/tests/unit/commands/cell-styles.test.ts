import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  activeCellStyleId,
  applyCellStyle,
  applyCellStyleByName,
  applyCellStyleToSelection,
  CELL_STYLES,
  createCellStyleFromActiveFormat,
  customCellStyleId,
  getCellStyle,
  listCustomCellStyles,
  mergeCellStylesFromWorkbook,
} from '../../../src/commands/cell-styles.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { fixedFormPolicy } from '../../../src/commands/interaction-policy.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/index.js';

const macStyleOracle = JSON.parse(
  readFileSync(
    resolve(dirname(fileURLToPath(import.meta.url)), '../../fixtures/excel-mac-style-groups.json'),
    'utf8',
  ),
) as {
  styles: readonly {
    galleryId: string | null;
    includedGroups: readonly string[];
  }[];
};

const namedStyleWorkbook = (): WorkbookHandle =>
  ({
    getNamedCellStyles: () => [
      {
        index: 0,
        name: 'Normal',
        xfId: 0,
        builtinId: 0,
        iLevel: 0,
        customBuiltin: false,
      },
      {
        index: 1,
        name: 'Imported Review',
        xfId: 2,
        builtinId: -1,
        iLevel: 0,
        customBuiltin: false,
      },
    ],
    getCellStyleXf: (xfId: number) =>
      xfId === 2
        ? {
            fontIndex: 1,
            fillIndex: 1,
            borderIndex: 0,
            numFmtId: 0,
            horizontalAlign: 2,
            verticalAlign: 1,
            wrapText: true,
          }
        : {
            fontIndex: 0,
            fillIndex: 0,
            borderIndex: 0,
            numFmtId: 0,
            horizontalAlign: 0,
            verticalAlign: 2,
            wrapText: false,
          },
    getFontRecord: (fontIndex: number) => ({
      name: 'Calibri',
      size: 11,
      bold: fontIndex === 1,
      italic: false,
      strike: false,
      underline: 0,
      colorArgb: fontIndex === 1 ? 0xff006100 : 0xff000000,
    }),
    getFillRecord: (fillIndex: number) => ({
      pattern: fillIndex === 1 ? 1 : 0,
      fgArgb: fillIndex === 1 ? 0xffc6efce : 0,
      bgArgb: 0,
    }),
    getBorderRecord: () => ({
      left: { style: 0, colorArgb: 0xff000000 },
      right: { style: 0, colorArgb: 0xff000000 },
      top: { style: 0, colorArgb: 0xff000000 },
      bottom: { style: 0, colorArgb: 0xff000000 },
      diagonal: { style: 0, colorArgb: 0xff000000 },
      diagonalUp: false,
      diagonalDown: false,
    }),
    getNumFmtCode: () => null,
  }) as unknown as WorkbookHandle;

describe('CELL_STYLES', () => {
  it('contains the spreadsheet-flavored presets', () => {
    const ids = CELL_STYLES.map((s) => s.id);
    expect(ids).toContain('normal');
    expect(ids).toContain('heading1');
    expect(ids).toContain('good');
    expect(ids).toContain('checkCell');
    expect(ids).toContain('explanatoryText');
    expect(ids).toContain('accent1');
    expect(ids).toContain('accent6_20');
    expect(ids).toContain('currency');
  });

  it('matches every observed Mac gallery include flag', () => {
    const observed = new Map(
      macStyleOracle.styles
        .filter((style): style is typeof style & { galleryId: string } => style.galleryId !== null)
        .map((style) => [style.galleryId, [...style.includedGroups].sort()]),
    );
    expect(observed.size).toBe(35);
    for (const style of CELL_STYLES) {
      expect([...(style.includedGroups ?? []).slice().sort()]).toEqual(observed.get(style.id));
    }
  });

  it('keeps native font discriminator bits in portable fallbacks', () => {
    expect(getCellStyle('title')?.format.bold).toBe(false);
    expect(getCellStyle('warning')?.format.italic).toBe(false);
    expect(getCellStyle('linkedCell')?.format.italic).toBe(false);
    expect(getCellStyle('calculation')?.format.italic).toBe(false);
    expect(getCellStyle('explanatoryText')?.builtinId).toBe(53);
    expect(getCellStyle('heading3')?.format).toMatchObject({ bold: true, fontSize: 11 });
    expect(getCellStyle('heading4')?.format).toMatchObject({
      bold: true,
      italic: false,
      fontSize: 11,
    });
  });
});

describe('getCellStyle', () => {
  it('returns the matching def', () => {
    expect(getCellStyle('good')?.format.fill).toBe('#c6efce');
    expect(getCellStyle('accent1')?.format.fill).toBe('#4472c4');
    expect(getCellStyle('accent4_20')?.format.fill).toBe('#fff2cc');
  });

  it('returns undefined for unknown ids', () => {
    // @ts-expect-error — testing invalid id rejection
    expect(getCellStyle('not-a-style')).toBeUndefined();
  });
});

describe('applyCellStyle', () => {
  it('writes the named-style fields onto every cell in range', () => {
    const store = createSpreadsheetStore();
    applyCellStyle(store, null, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }, 'good');
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.fill).toBe('#c6efce');
    expect(fmt?.color).toBe('#006100');
    expect(fmt?.cellStyle).toBe('good');
    const fmtCorner = store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 }));
    expect(fmtCorner?.fill).toBe('#c6efce');
    expect(fmtCorner?.cellStyle).toBe('good');
  });

  it('clears every format field for the "normal" preset', () => {
    const store = createSpreadsheetStore();
    applyCellStyle(store, null, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, 'good');
    applyCellStyle(store, null, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, 'normal');
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    // Either the entry is gone, or all visible style fields are undefined.
    if (fmt) {
      expect(fmt.fill).toBeUndefined();
      expect(fmt.color).toBeUndefined();
      expect(fmt.bold).toBeUndefined();
      expect(fmt.cellStyle).toBeUndefined();
    }
  });

  it('applies Excel-style accent and explanatory presets', () => {
    const store = createSpreadsheetStore();
    applyCellStyle(store, null, { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 }, 'accent5_20');
    let fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 2 }));
    expect(fmt).toMatchObject({ color: '#1f4e79', fill: '#ddebf7' });

    applyCellStyle(store, null, { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 }, 'explanatoryText');
    fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 2 }));
    expect(fmt).toMatchObject({ color: '#7f7f7f', italic: true });
  });

  it('no-ops on unknown style id', () => {
    const store = createSpreadsheetStore();
    const before = store.getState();
    // @ts-expect-error — testing invalid id rejection
    applyCellStyle(store, null, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, 'not-a-style');
    expect(store.getState()).toBe(before);
  });

  it('creates a named style from the active cell format and records history', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        bold: true,
        fill: '#c6efce',
        color: '#006100',
      },
    );

    expect(
      createCellStyleFromActiveFormat(
        store,
        history,
        { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
        'Review OK',
        { include: { fill: false } },
      ),
    ).toBe(true);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 })),
    ).toMatchObject({
      cellStyle: 'Review OK',
      bold: true,
      color: '#006100',
    });
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 }))?.fill,
    ).toBeUndefined();
    expect(listCustomCellStyles(store.getState())).toMatchObject([
      {
        id: customCellStyleId('Review OK'),
        label: 'Review OK',
        format: {
          bold: true,
          color: '#006100',
        },
      },
    ]);
    expect(
      applyCellStyleByName(
        store,
        history,
        { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
        customCellStyleId('Review OK'),
      ),
    ).toBe(true);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 2 })),
    ).toMatchObject({
      cellStyle: 'Review OK',
      bold: true,
      color: '#006100',
    });
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 2 }))?.fill,
    ).toBeUndefined();
    expect(history.undo()).toBe(true);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 2 })),
    ).toBeUndefined();
    expect(history.undo()).toBe(true);
    expect(listCustomCellStyles(store.getState())).toEqual([]);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 })),
    ).toBeUndefined();
  });

  it('merges visible workbook named styles into the session custom style registry', () => {
    const store = createSpreadsheetStore();
    const history = new History();

    expect(mergeCellStylesFromWorkbook(store, history, namedStyleWorkbook())).toEqual({
      imported: 1,
      skipped: 1,
    });
    expect(listCustomCellStyles(store.getState())).toMatchObject([
      {
        id: customCellStyleId('Imported Review'),
        label: 'Imported Review',
        format: {
          bold: true,
          color: '#006100',
          fill: '#c6efce',
          align: 'center',
          vAlign: 'middle',
          wrap: true,
        },
      },
    ]);
    expect(
      applyCellStyleByName(
        store,
        history,
        { sheet: 0, r0: 3, c0: 3, r1: 3, c1: 3 },
        customCellStyleId('Imported Review'),
      ),
    ).toBe(true);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 3, col: 3 })),
    ).toMatchObject({
      cellStyle: 'Imported Review',
      bold: true,
      fill: '#c6efce',
    });
    expect(history.undo()).toBe(true);
    expect(history.undo()).toBe(true);
    expect(listCustomCellStyles(store.getState())).toEqual([]);
  });

  it('repeats a gallery style onto whatever is selected when F4 fires', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    applyCellStyle(store, history, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, 'good');

    mutators.setRange(store, { sheet: 0, r0: 4, c0: 2, r1: 4, c1: 2 });
    expect(history.repeatLast()).toBe(true);

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 4, col: 2 })),
    ).toMatchObject({ cellStyle: 'good' });
  });

  it('repeats a custom style by name rather than redefining it', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.setRangeFormat(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, { bold: true });
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    expect(
      createCellStyleFromActiveFormat(
        store,
        history,
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        'Brand',
      ),
    ).toBe(true);

    mutators.setRange(store, { sheet: 0, r0: 2, c0: 0, r1: 2, c1: 0 });
    expect(history.repeatLast()).toBe(true);

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 0 })),
    ).toMatchObject({ cellStyle: 'Brand', bold: true });
    expect(listCustomCellStyles(store.getState())).toHaveLength(1);
  });

  it('applies a built-in style across the complete primary-plus-extra union with group replacement', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
        extraRanges: [
          { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 2 },
          { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
        ],
      },
    }));
    for (const col of [0, 1, 2]) {
      mutators.setCellFormat(
        store,
        { sheet: 0, row: 0, col },
        {
          bold: true,
          italic: true,
          fill: '#123456',
          numFmt: { kind: 'percent', decimals: 1 },
          locked: false,
          comment: `keep-${col}`,
          hyperlink: `https://example.test/${col}`,
        },
      );
    }
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 2, col: 2 },
      {
        bold: true,
        fill: '#123456',
        numFmt: { kind: 'percent', decimals: 1 },
        locked: false,
        comment: 'keep-hole',
      },
    );

    expect(
      applyCellStyleToSelection(store, history, 'good', {
        origin: 'instanceApi',
        commandId: 'cellStyles',
      }),
    ).toBe(true);

    for (const addr of [
      { sheet: 0, row: 0, col: 0 },
      { sheet: 0, row: 0, col: 1 },
      { sheet: 0, row: 0, col: 2 },
    ]) {
      expect(store.getState().format.formats.get(addrKey(addr))).toMatchObject({
        color: '#006100',
        fill: '#c6efce',
        numFmt: { kind: 'percent', decimals: 1 },
        locked: false,
      });
      expect(store.getState().format.formats.get(addrKey(addr))?.bold).toBeUndefined();
      expect(store.getState().format.formats.get(addrKey(addr))?.comment).toMatch(/keep/);
    }
    expect(store.getState().format.formats.has('0:1:1')).toBe(false);
    expect(history.undo()).toBe(true);
    expect(store.getState().format.formats.get('0:0:0')?.bold).toBe(true);
    expect(history.redo()).toBe(true);
    expect(store.getState().format.formats.get('0:0:0')?.fill).toBe('#c6efce');
  });

  it('resolves the active merge anchor and preserves unresolved style keys', () => {
    const store = createSpreadsheetStore();
    mutators.mergeRange(store, { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 3 });
    mutators.setCellFormat(store, { sheet: 0, row: 2, col: 2 }, { cellStyle: 'good' });
    mutators.setActive(store, { sheet: 0, row: 2, col: 3 });
    expect(activeCellStyleId(store.getState())).toBe('good');

    mutators.setCellFormat(store, { sheet: 0, row: 2, col: 2 }, { cellStyle: 'Imported Review' });
    expect(activeCellStyleId(store.getState())).toBe('Imported Review');
    mutators.setCellFormat(store, { sheet: 0, row: 2, col: 2 }, { cellStyle: undefined });
    expect(activeCellStyleId(store.getState())).toBe('normal');
  });

  it('uses the observed group boundaries for Note and Heading 3', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setCellFormat(store, addr, {
      bold: true,
      color: '#123456',
      fill: '#abcdef',
      numFmt: { kind: 'percent', decimals: 2 },
      locked: false,
      borders: {
        top: { style: 'thick', color: '#ff0000' },
        bottom: { style: 'thin', color: '#00ff00' },
      },
    });

    expect(applyCellStyleToSelection(store, history, 'note')).toBe(true);
    let format = store.getState().format.formats.get(addrKey(addr));
    expect(format).toMatchObject({
      bold: true,
      color: '#123456',
      fill: '#ffffcc',
      numFmt: { kind: 'percent', decimals: 2 },
      locked: false,
    });
    expect(format?.borders).toBeUndefined();

    expect(applyCellStyleToSelection(store, history, 'heading3')).toBe(true);
    format = store.getState().format.formats.get(addrKey(addr));
    expect(format).toMatchObject({ bold: true, color: '#1f4e79', fill: '#ffffcc' });
    expect(format?.borders).toBeUndefined();
  });

  it('uses the same replacement groups through the explicit range API', () => {
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 1, col: 1 };
    mutators.setCellFormat(store, addr, {
      bold: true,
      color: '#123456',
      fill: '#abcdef',
      numFmt: { kind: 'currency', decimals: 2, symbol: '$' },
      borders: { top: { style: 'thick' } },
    });

    applyCellStyle(store, null, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 }, 'note');
    const format = store.getState().format.formats.get(addrKey(addr));
    expect(format).toMatchObject({
      bold: true,
      color: '#123456',
      fill: '#ffffcc',
      numFmt: { kind: 'currency', decimals: 2, symbol: '$' },
      cellStyle: 'note',
    });
    expect(format?.borders).toBeUndefined();
  });

  it('keeps explicit range styling scoped and respects protection and the materialization cap', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.upsertCustomCellStyle(store, {
      id: 'custom:Scoped',
      label: 'Scoped',
      format: { bold: true },
    });
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 1 }, { fill: '#abcdef' });
    store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        extraRanges: [{ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }],
      },
    }));

    expect(
      applyCellStyleByName(
        store,
        history,
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        customCellStyleId('Scoped'),
      ),
    ).toBe(true);
    expect(store.getState().format.formats.get('0:0:0')?.bold).toBe(true);
    expect(store.getState().format.formats.get('0:0:1')?.fill).toBe('#abcdef');

    const locked = createSpreadsheetStore();
    const lockedHistory = new History();
    mutators.upsertCustomCellStyle(locked, {
      id: 'custom:Scoped',
      label: 'Scoped',
      format: { bold: true },
    });
    mutators.setSheetProtected(locked, 0, true);
    expect(
      applyCellStyleByName(
        locked,
        lockedHistory,
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        customCellStyleId('Scoped'),
      ),
    ).toBe(false);
    expect(locked.getState().format.formats.size).toBe(0);
    expect(lockedHistory.repeatLast()).toBe(false);

    const huge = createSpreadsheetStore();
    expect(
      applyCellStyle(huge, null, { sheet: 0, r0: 0, c0: 0, r1: 100_000, c1: 1 }, 'good'),
    ).toBeUndefined();
    expect(huge.getState().format.formats.size).toBe(0);
  });

  it('authorizes the supplied range independently from the live selection', async () => {
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.upsertCustomCellStyle(store, {
      id: 'custom:ScopedPolicy',
      label: 'ScopedPolicy',
      format: { bold: true },
    });
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    mutators.addExtraRange(store, { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 });
    const liveSelection = store.getState().selection;
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    const unregister = registerInteractionController(store, controller);
    const policy = fixedFormPolicy({ ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }] });
    controller.setPolicy({ ...policy, operations: { ...policy.operations, format: true } });
    try {
      expect(
        applyCellStyleByName(
          store,
          history,
          { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
          customCellStyleId('ScopedPolicy'),
        ),
      ).toBe(false);
      expect(store.getState().format.formats.size).toBe(0);
      expect(history.repeatLast()).toBe(false);

      const targetPolicy = fixedFormPolicy({
        ranges: [{ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }],
      });
      controller.setPolicy({
        ...targetPolicy,
        operations: { ...targetPolicy.operations, format: true },
      });
      expect(
        applyCellStyleByName(
          store,
          history,
          { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
          customCellStyleId('ScopedPolicy'),
        ),
      ).toBe(true);
      expect(store.getState().format.formats.get('0:0:1')?.bold).toBe(true);
      expect(store.getState().format.formats.has('0:0:2')).toBe(false);
      expect(store.getState().selection).toEqual(liveSelection);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  it('writes only unlocked cells in one protected-range history entry', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.upsertCustomCellStyle(store, {
      id: 'custom:ProtectedMixed',
      label: 'ProtectedMixed',
      format: { bold: true },
    });
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { locked: false });
    mutators.setSheetProtected(store, 0, true);
    expect(
      applyCellStyleByName(
        store,
        history,
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
        customCellStyleId('ProtectedMixed'),
      ),
    ).toBe(true);
    expect(store.getState().format.formats.get('0:0:0')).toMatchObject({
      locked: false,
      bold: true,
    });
    expect(store.getState().format.formats.get('0:0:1')).toBeUndefined();
    expect(history.undo()).toBe(true);
    expect(store.getState().format.formats.get('0:0:0')).toMatchObject({ locked: false });
    expect(store.getState().format.formats.get('0:0:0')?.bold).toBeUndefined();
  });

  it('rejects malformed, foreign-selection, and partial-merge plans without replacing repeat', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.upsertCustomCellStyle(store, {
      id: 'custom:StableRepeat',
      label: 'StableRepeat',
      format: { italic: true },
    });
    expect(
      applyCellStyleByName(
        store,
        history,
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        customCellStyleId('StableRepeat'),
      ),
    ).toBe(true);
    mutators.setRange(store, { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 });
    const selectionBefore = store.getState().selection;
    expect(
      applyCellStyleByName(
        store,
        history,
        { sheet: 0, r0: 2, c0: 0, r1: 1, c1: 0 },
        customCellStyleId('StableRepeat'),
      ),
    ).toBe(false);
    expect(store.getState().selection).toEqual(selectionBefore);
    expect(history.repeatLast()).toBe(true);
    expect(store.getState().format.formats.get('0:1:0')?.italic).toBe(true);

    const foreignSelection = createSpreadsheetStore();
    mutators.upsertCustomCellStyle(foreignSelection, {
      id: 'custom:StableRepeat',
      label: 'StableRepeat',
      format: { italic: true },
    });
    foreignSelection.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        extraRanges: [{ sheet: 1, r0: 0, c0: 0, r1: 0, c1: 0 }],
      },
    }));
    expect(
      applyCellStyleToSelection(foreignSelection, new History(), customCellStyleId('StableRepeat')),
    ).toBe(false);
    expect(foreignSelection.getState().format.formats.size).toBe(0);

    const merged = createSpreadsheetStore();
    mutators.upsertCustomCellStyle(merged, {
      id: 'custom:StableRepeat',
      label: 'StableRepeat',
      format: { italic: true },
    });
    mutators.mergeRange(merged, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    expect(
      applyCellStyleByName(
        merged,
        new History(),
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
        customCellStyleId('StableRepeat'),
      ),
    ).toBe(false);
    expect(merged.getState().format.formats.size).toBe(0);
  });

  it('returns false for exactly 100001 cells without replacing history or repeat', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.upsertCustomCellStyle(store, {
      id: 'custom:TooWide',
      label: 'TooWide',
      format: { bold: true },
    });
    expect(
      applyCellStyleByName(
        store,
        history,
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        customCellStyleId('TooWide'),
      ),
    ).toBe(true);
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 });
    expect(
      applyCellStyleByName(
        store,
        history,
        { sheet: 0, r0: 0, c0: 0, r1: 100_000, c1: 0 },
        customCellStyleId('TooWide'),
      ),
    ).toBe(false);
    expect(store.getState().format.formats.get('0:0:0')?.bold).toBe(true);
    expect(history.canUndo()).toBe(true);
    expect(history.repeatLast()).toBe(true);
    expect(store.getState().format.formats.get('0:0:1')?.bold).toBe(true);
  });

  it('accepts a complete merge and rejects a partial merge through explicit APIs', () => {
    const complete = createSpreadsheetStore();
    const history = new History();
    const merge = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    mutators.mergeRange(complete, merge);
    expect(applyCellStyleByName(complete, history, merge, customCellStyleId('Missing'))).toBe(
      false,
    );
    mutators.upsertCustomCellStyle(complete, {
      id: 'custom:Merged',
      label: 'Merged',
      format: { fill: '#c6efce' },
    });
    expect(applyCellStyleByName(complete, history, merge, customCellStyleId('Merged'))).toBe(true);
    expect(complete.getState().format.formats.get('0:0:0')?.fill).toBe('#c6efce');
    expect(complete.getState().format.formats.get('0:1:1')?.fill).toBe('#c6efce');

    const partial = createSpreadsheetStore();
    mutators.mergeRange(partial, merge);
    mutators.upsertCustomCellStyle(partial, {
      id: 'custom:Partial',
      label: 'Partial',
      format: { fill: '#ffc7ce' },
    });
    expect(
      applyCellStyleByName(
        partial,
        new History(),
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
        customCellStyleId('Partial'),
      ),
    ).toBe(false);
    expect(partial.getState().format.formats.size).toBe(0);
  });
});
