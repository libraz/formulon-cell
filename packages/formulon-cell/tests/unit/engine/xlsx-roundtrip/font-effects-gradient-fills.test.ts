import { describe, expect, it } from 'vitest';
import {
  hydrateCellFormatsFromEngine,
  syncCellFormatsToEngine,
} from '../../../../src/engine/cell-format-sync.js';
import type { ColorSpec, FillRecord } from '../../../../src/engine/types.js';
import { addrKey, WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import {
  type CellFormat,
  createSpreadsheetStore,
  type SpreadsheetStore,
} from '../../../../src/store/store.js';

import { canLoadWasm } from './fixtures.js';

const A1 = { sheet: 0, row: 0, col: 0 };
const B1 = { sheet: 0, row: 0, col: 1 };

const literal = (rgb: number): ColorSpec => ({ kind: 1, rgb, theme: 0, tint: 0, indexed: 0 });
const themed = (theme: number, tint: number): ColorSpec => ({
  kind: 2,
  rgb: 0,
  theme,
  tint,
  indexed: 0,
});

const emptyBorder = {
  left: { style: 0, colorArgb: 0 },
  right: { style: 0, colorArgb: 0 },
  top: { style: 0, colorArgb: 0 },
  bottom: { style: 0, colorArgb: 0 },
  diagonal: { style: 0, colorArgb: 0 },
  diagonalUp: false,
  diagonalDown: false,
};

const setFormat = (store: SpreadsheetStore, key: string, format: CellFormat): void => {
  store.setState((s) => {
    const formats = new Map(s.format.formats);
    formats.set(key, format);
    return { ...s, format: { ...s.format, formats } };
  });
};

const formatAt = (store: SpreadsheetStore, addr: typeof A1): CellFormat | undefined =>
  store.getState().format.formats.get(addrKey(addr));

const fontAt = (wb: WorkbookHandle, addr: typeof A1) => {
  const xf = wb.getCellXf(wb.getCellXfIndex(addr.sheet, addr.row, addr.col) ?? -1);
  return xf ? wb.getFontRecord(xf.fontIndex) : null;
};

const fillAt = (wb: WorkbookHandle, addr: typeof A1): FillRecord | null => {
  const xf = wb.getCellXf(wb.getCellXfIndex(addr.sheet, addr.row, addr.col) ?? -1);
  return xf ? wb.getFillRecord(xf.fillIndex) : null;
};

const reload = async (wb: WorkbookHandle): Promise<WorkbookHandle> =>
  WorkbookHandle.loadBytes(wb.save());

/** A workbook whose A1 carries a raw font / fill only a file can author. */
const seedRawCell = async (
  font: Partial<Parameters<WorkbookHandle['addFontRecord']>[0]>,
  fill?: FillRecord,
): Promise<WorkbookHandle> => {
  const first = await WorkbookHandle.createDefault();
  first.setText(A1, 'x');
  first.setText(B1, 'y');
  const fontIndex = first.addFontRecord({
    name: 'Arial',
    size: 11,
    bold: false,
    italic: false,
    strike: false,
    underline: 0,
    colorArgb: 0xff000000,
    ...font,
  });
  const fillIndex = first.addFillRecord(fill ?? { pattern: 0, fgArgb: 0, bgArgb: 0 });
  const borderIndex = first.addBorderRecord(emptyBorder);
  const xf = first.addXfRecord({
    fontIndex,
    fillIndex,
    borderIndex,
    numFmtId: 0,
    horizontalAlign: 0,
    verticalAlign: 2,
    wrapText: false,
  });
  first.setCellXfIndex(0, 0, 0, xf);
  try {
    return await reload(first);
  } finally {
    first.dispose();
  }
};

const GRADIENT: FillRecord = {
  pattern: 0,
  fgArgb: 0,
  bgArgb: 0,
  gradient: {
    type: 0,
    degree: 90,
    left: 0,
    right: 0,
    top: 0,
    bottom: 0,
    stops: [
      { position: 0, colorArgb: 0xffff0000, color: literal(0xffff0000) },
      { position: 1, colorArgb: 0xff4472c4, color: themed(4, 0.4) },
    ],
  },
};

describe.skipIf(!canLoadWasm())('font effects and gradient fills through real xlsx', () => {
  it('saves and reloads authored outline, shadow, condense and extend', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    try {
      wb.setText(A1, 'x');
      setFormat(store, addrKey(A1), {
        fontOutline: true,
        fontShadow: true,
        fontCondense: true,
        fontExtend: true,
      });
      syncCellFormatsToEngine(wb, store, 0);
      const reloaded = await reload(wb);
      try {
        const next = createSpreadsheetStore();
        hydrateCellFormatsFromEngine(reloaded, next, 0);
        expect(formatAt(next, A1)).toMatchObject({
          fontOutline: true,
          fontShadow: true,
          fontCondense: true,
          fontExtend: true,
        });
      } finally {
        reloaded.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('keeps loaded font effects when the cell is reformatted or its format is copied', async () => {
    const wb = await seedRawCell({
      outline: true,
      hasOutline: true,
      condense: true,
      hasCondense: true,
      extend: true,
      hasExtend: true,
    });
    const store = createSpreadsheetStore();
    try {
      hydrateCellFormatsFromEngine(wb, store, 0);
      const loaded = formatAt(store, A1) ?? {};
      setFormat(store, addrKey(A1), { ...loaded, italic: true });
      setFormat(store, addrKey(B1), { ...loaded, bold: true });
      syncCellFormatsToEngine(wb, store, 0);
      const reloaded = await reload(wb);
      try {
        for (const addr of [A1, B1]) {
          expect(fontAt(reloaded, addr)).toMatchObject({
            outline: true,
            condense: true,
            extend: true,
            shadow: false,
          });
        }
        expect(fontAt(reloaded, A1)?.italic).toBe(true);
        expect(fontAt(reloaded, B1)?.bold).toBe(true);
      } finally {
        reloaded.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('switching a loaded font effect off writes it off', async () => {
    const wb = await seedRawCell({ shadow: true, hasShadow: true });
    const store = createSpreadsheetStore();
    try {
      hydrateCellFormatsFromEngine(wb, store, 0);
      expect(formatAt(store, A1)?.fontShadow).toBe(true);
      setFormat(store, addrKey(A1), { ...(formatAt(store, A1) ?? {}), fontShadow: false });
      syncCellFormatsToEngine(wb, store, 0);
      expect(fontAt(wb, A1)?.shadow).toBe(false);
    } finally {
      wb.dispose();
    }
  });

  it('hydrates a loaded gradient fill and keeps it, theme stop included, through a reformat', async () => {
    const wb = await seedRawCell({}, GRADIENT);
    const store = createSpreadsheetStore();
    try {
      hydrateCellFormatsFromEngine(wb, store, 0);
      const loaded = formatAt(store, A1) ?? {};
      expect(loaded.fillGradient).toMatchObject({
        type: 'linear',
        degree: 90,
        stops: [
          { position: 0, color: '#ff0000' },
          { position: 1, color: expect.stringMatching(/^#[0-9a-f]{6}$/) },
        ],
      });
      setFormat(store, addrKey(A1), { ...loaded, bold: true });
      setFormat(store, addrKey(B1), { ...loaded });
      syncCellFormatsToEngine(wb, store, 0);
      const reloaded = await reload(wb);
      try {
        const kept = fillAt(reloaded, A1)?.gradient;
        expect(kept?.degree).toBe(90);
        expect(kept?.stops[1]?.color).toMatchObject({ kind: 2, theme: 4 });
        expect(kept?.stops[1]?.color?.tint).toBeCloseTo(0.4, 5);
        expect(fontAt(reloaded, A1)?.bold).toBe(true);
        const copied = fillAt(reloaded, B1)?.gradient;
        expect(copied?.type).toBe(0);
        expect(copied?.stops.map((s) => s.position)).toEqual([0, 1]);
        expect(copied?.stops[0]?.colorArgb).toBe(0xffff0000);
      } finally {
        reloaded.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('saves and reloads an authored path gradient', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const fillGradient: NonNullable<CellFormat['fillGradient']> = {
      type: 'path',
      degree: 0,
      left: 0.5,
      right: 0.5,
      top: 0.5,
      bottom: 0.5,
      stops: [
        { position: 0, color: '#ffffff' },
        { position: 1, color: '#4472c4' },
      ],
    };
    try {
      wb.setText(A1, 'x');
      setFormat(store, addrKey(A1), { fillGradient });
      syncCellFormatsToEngine(wb, store, 0);
      const reloaded = await reload(wb);
      try {
        const next = createSpreadsheetStore();
        hydrateCellFormatsFromEngine(reloaded, next, 0);
        expect(formatAt(next, A1)?.fillGradient).toEqual(fillGradient);
      } finally {
        reloaded.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('a solid fill replaces a loaded gradient', async () => {
    const wb = await seedRawCell({}, GRADIENT);
    const store = createSpreadsheetStore();
    try {
      hydrateCellFormatsFromEngine(wb, store, 0);
      const { fillGradient: _drop, ...rest } = formatAt(store, A1) ?? {};
      setFormat(store, addrKey(A1), { ...rest, fill: '#00ff00' });
      syncCellFormatsToEngine(wb, store, 0);
      const fill = fillAt(wb, A1);
      expect(fill?.gradient).toBeUndefined();
      expect(fill?.pattern).toBe(1);
    } finally {
      wb.dispose();
    }
  });
});
