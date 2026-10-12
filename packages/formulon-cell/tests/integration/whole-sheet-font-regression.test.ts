import { describe, expect, it } from 'vitest';
import { clearVisualFormat, setFont } from '../../src/commands/format.js';
import { History } from '../../src/commands/history.js';
import { moveSheet, removeSheet } from '../../src/commands/sheet-mutate.js';
import { recordFormatChange } from '../../src/commands/slice-history.js';
import { MAX_COL, MAX_ROW } from '../../src/engine/address.js';
import { flushFormatToEngine } from '../../src/engine/cell-format-sync.js';
import type { CellXf, FontRecord } from '../../src/engine/types.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { mutators } from '../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/mount.js';

const target = { sheet: 0, row: 0, col: 0 };
const other = { sheet: 1, row: 0, col: 0 };
const far = { sheet: 0, row: MAX_ROW, col: MAX_COL };

function fontAt(workbook: WorkbookHandle, addr: typeof target) {
  const style = workbook.getEffectiveStyle(addr.sheet, addr.row, addr.col);
  const xf = style && workbook.getCellXf(style.xfIndex);
  const font = xf && workbook.getFontRecord(xf.fontIndex);
  if (!font) throw new Error(`missing font ${addr.sheet}:${addr.row}:${addr.col}`);
  return font;
}

function rawAt(workbook: WorkbookHandle, addr: typeof far = far): { xf: CellXf; font: FontRecord } {
  const style = workbook.getEffectiveStyle(addr.sheet, addr.row, addr.col);
  const xf = style && workbook.getCellXf(style.xfIndex);
  const font = xf && workbook.getFontRecord(xf.fontIndex);
  if (!xf || !font)
    throw new Error(`missing effective style ${addr.sheet}:${addr.row}:${addr.col}`);
  return { xf, font };
}

function installImportedAxisFont(workbook: WorkbookHandle): FontRecord {
  const baseFont = workbook.getFontRecord(0);
  const baseXf = workbook.getCellXf(0);
  if (!baseFont || !baseXf) throw new Error('missing workbook defaults');
  const imported: FontRecord = {
    ...baseFont,
    name: 'Consolas',
    size: 15,
    bold: true,
    hasBold: true,
    charset: 128,
    hasCharset: true,
  };
  const xfIndex = workbook.addXfRecord({
    ...baseXf,
    fontIndex: workbook.addFontRecord(imported),
    numFmtId: 14,
    horizontalAlign: 2,
    verticalAlign: 1,
    wrapText: true,
    hasAlignment: true,
    hasHorizontalAlign: true,
    hasVerticalAlign: true,
    hasWrapText: true,
    applyNumberFormat: true,
    applyAlignment: true,
    applyProtection: true,
    hasProtection: true,
    locked: false,
    hidden: true,
    quotePrefix: true,
  });
  if (!Number.isInteger(xfIndex) || xfIndex < 0) throw new Error('failed to add imported XF');
  expect(workbook.setColumnStyle(0, 0, MAX_COL, xfIndex)).toBe(true);
  return imported;
}

function seedDistinctOtherCell(mounted: MountedStubSheet): void {
  mutators.setCellFormat(mounted.instance.store, other, {
    fontFamily: 'Consolas',
    fontSize: 20,
    bold: true,
    fill: '#fff2cc',
  });
  expect(fontAt(mounted.workbook, other)).toMatchObject({
    name: 'Consolas',
    size: 15,
    bold: true,
  });
}

async function mountPopulatedWorkbook(): Promise<MountedStubSheet> {
  const workbook = await WorkbookHandle.createDefault({ locale: 'ja' });
  expect(workbook.isStub).toBe(false);
  workbook.addSheet('Other');
  workbook.setText(target, 'target');
  workbook.setText(other, 'other');
  return mountStubSheet({ workbook, locale: 'ja' });
}

describe('whole-sheet font structure remap regression', () => {
  it('moves populated sparse formats and high-water keys without resetting either XF', async () => {
    const mounted = await mountPopulatedWorkbook();
    try {
      seedDistinctOtherCell(mounted);
      mutators.selectAll(mounted.instance.store);
      expect(
        setFont(mounted.instance.store.getState(), mounted.instance.store, { fontSize: 24 }),
      ).toBe(true);
      expect(fontAt(mounted.workbook, target)).toMatchObject({ size: 18 });

      expect(mounted.workbook.moveSheet(0, 1)).toBe(true);
      expect(fontAt(mounted.workbook, { ...target, sheet: 0 })).toMatchObject({
        name: 'Consolas',
        size: 15,
        bold: true,
      });
      expect(fontAt(mounted.workbook, { ...target, sheet: 1 })).toMatchObject({ size: 18 });
      const formats = mounted.instance.store.getState().format.formats;
      expect(formats.get('0:0:0')).toMatchObject({ fontFamily: 'Consolas', fontSize: 20 });
      expect(formats.get('1:0:0')).toMatchObject({ fontSize: 24 });
      expect([
        ...(mounted.instance.store.getState().format.sheetFontOverrides?.keys() ?? []),
      ]).toEqual([1]);

      flushFormatToEngine(mounted.workbook, mounted.instance.store, 0);
      flushFormatToEngine(mounted.workbook, mounted.instance.store, 1);
      expect(fontAt(mounted.workbook, { ...target, sheet: 0 })).toMatchObject({
        name: 'Consolas',
        size: 15,
        bold: true,
      });
      expect(fontAt(mounted.workbook, { ...target, sheet: 1 })).toMatchObject({ size: 18 });

      const reloaded = await WorkbookHandle.loadBytes(mounted.workbook.save());
      try {
        expect(fontAt(reloaded, { ...target, sheet: 0 })).toMatchObject({
          name: 'Consolas',
          size: 15,
          bold: true,
        });
        expect(fontAt(reloaded, { ...target, sheet: 1 })).toMatchObject({ size: 18 });
      } finally {
        reloaded.dispose();
      }
    } finally {
      mounted.dispose();
      mounted.workbook.dispose();
    }
  });

  it('removes populated sparse formats without resetting the surviving sheet XF', async () => {
    const mounted = await mountPopulatedWorkbook();
    try {
      seedDistinctOtherCell(mounted);
      mutators.selectAll(mounted.instance.store);
      expect(
        setFont(mounted.instance.store.getState(), mounted.instance.store, { fontSize: 24 }),
      ).toBe(true);

      expect(mounted.workbook.removeSheet(0)).toBe(true);
      expect(fontAt(mounted.workbook, { ...target, sheet: 0 })).toMatchObject({
        name: 'Consolas',
        size: 15,
        bold: true,
      });
      expect(mounted.instance.store.getState().format.formats.get('0:0:0')).toMatchObject({
        fontFamily: 'Consolas',
        fontSize: 20,
      });
      expect(mounted.instance.store.getState().format.sheetFontOverrides?.size ?? 0).toBe(0);
      flushFormatToEngine(mounted.workbook, mounted.instance.store, 0);
      expect(fontAt(mounted.workbook, { ...target, sheet: 0 })).toMatchObject({
        name: 'Consolas',
        size: 15,
        bold: true,
      });

      const reloaded = await WorkbookHandle.loadBytes(mounted.workbook.save());
      try {
        expect(fontAt(reloaded, { ...target, sheet: 0 })).toMatchObject({
          name: 'Consolas',
          size: 15,
          bold: true,
        });
      } finally {
        reloaded.dispose();
      }
    } finally {
      mounted.dispose();
      mounted.workbook.dispose();
    }
  });

  it('keeps distinct populated formats through the command move route', async () => {
    const mounted = await mountPopulatedWorkbook();
    try {
      seedDistinctOtherCell(mounted);
      mutators.selectAll(mounted.instance.store);
      expect(
        setFont(mounted.instance.store.getState(), mounted.instance.store, { fontSize: 24 }),
      ).toBe(true);
      expect(moveSheet(mounted.instance.store, mounted.workbook, 0, 1)).toBe(true);
      flushFormatToEngine(mounted.workbook, mounted.instance.store, 0);
      flushFormatToEngine(mounted.workbook, mounted.instance.store, 1);
      expect(fontAt(mounted.workbook, { ...target, sheet: 0 })).toMatchObject({
        name: 'Consolas',
        size: 15,
        bold: true,
      });
      expect(fontAt(mounted.workbook, { ...target, sheet: 1 })).toMatchObject({ size: 18 });
    } finally {
      mounted.dispose();
      mounted.workbook.dispose();
    }
  });

  it('keeps distinct populated formats through the command remove route', async () => {
    const mounted = await mountPopulatedWorkbook();
    try {
      seedDistinctOtherCell(mounted);
      mutators.selectAll(mounted.instance.store);
      expect(
        setFont(mounted.instance.store.getState(), mounted.instance.store, { fontSize: 24 }),
      ).toBe(true);
      expect(removeSheet(mounted.instance.store, mounted.workbook, 0)).toBe(true);
      flushFormatToEngine(mounted.workbook, mounted.instance.store, 0);
      expect(fontAt(mounted.workbook, { ...target, sheet: 0 })).toMatchObject({
        name: 'Consolas',
        size: 15,
        bold: true,
      });
    } finally {
      mounted.dispose();
      mounted.workbook.dispose();
    }
  });

  it('clears a loaded axis through undo/redo and preserves the cleared result after reload', async () => {
    const source = await WorkbookHandle.createDefault({ locale: 'ja' });
    installImportedAxisFont(source);
    const loaded = await WorkbookHandle.loadBytes(source.save());
    source.dispose();
    const mounted = await mountStubSheet({ workbook: loaded, locale: 'ja' });
    const history = new History();
    try {
      const defaultFont = loaded.getFontRecord(0);
      if (!defaultFont) throw new Error('missing default font');
      expect(rawAt(loaded).font).toMatchObject({ name: 'Consolas', size: 15, bold: true });
      mutators.selectAll(mounted.instance.store);
      recordFormatChange(history, mounted.instance.store, () => {
        clearVisualFormat(mounted.instance.store.getState(), mounted.instance.store);
      });
      expect(rawAt(loaded).font).toMatchObject({
        name: defaultFont.name,
        size: defaultFont.size,
        bold: false,
      });
      expect(history.undo()).toBe(true);
      expect(rawAt(loaded).font).toMatchObject({ name: 'Consolas', size: 15, bold: true });
      expect(history.redo()).toBe(true);
      expect(rawAt(loaded).font).toMatchObject({
        name: defaultFont.name,
        size: defaultFont.size,
        bold: false,
      });
      const reloaded = await WorkbookHandle.loadBytes(loaded.save());
      try {
        expect(rawAt(reloaded).font).toMatchObject({
          name: defaultFont.name,
          size: defaultFont.size,
          bold: false,
        });
      } finally {
        reloaded.dispose();
      }
    } finally {
      mounted.dispose();
      loaded.dispose();
    }
  });

  it('keeps axis protection and quotePrefix while clear then font removes old visuals', async () => {
    const source = await WorkbookHandle.createDefault({ locale: 'ja' });
    installImportedAxisFont(source);
    const loaded = await WorkbookHandle.loadBytes(source.save());
    source.dispose();
    const mounted = await mountStubSheet({ workbook: loaded, locale: 'ja' });
    try {
      const defaultXf = loaded.getCellXf(0);
      const defaultFont = loaded.getFontRecord(0);
      if (!defaultXf || !defaultFont) throw new Error('missing workbook defaults');
      mutators.selectAll(mounted.instance.store);
      clearVisualFormat(mounted.instance.store.getState(), mounted.instance.store);
      expect(rawAt(loaded).font).toMatchObject({
        name: defaultFont.name,
        size: defaultFont.size,
        bold: false,
      });
      expect(rawAt(loaded).xf).toMatchObject({
        fillIndex: defaultXf.fillIndex,
        borderIndex: defaultXf.borderIndex,
        numFmtId: defaultXf.numFmtId,
        horizontalAlign: defaultXf.horizontalAlign,
        verticalAlign: defaultXf.verticalAlign,
        wrapText: defaultXf.wrapText,
        applyProtection: true,
        hasProtection: true,
        locked: false,
        hidden: true,
        quotePrefix: true,
      });
      expect(
        setFont(mounted.instance.store.getState(), mounted.instance.store, {
          fontFamily: 'Arial',
        }),
      ).toBe(true);
      expect(rawAt(loaded).font).toMatchObject({
        name: 'Arial',
        size: defaultFont.size,
        bold: false,
      });
      expect(rawAt(loaded).xf).toMatchObject({
        fillIndex: defaultXf.fillIndex,
        borderIndex: defaultXf.borderIndex,
        numFmtId: defaultXf.numFmtId,
        horizontalAlign: defaultXf.horizontalAlign,
        verticalAlign: defaultXf.verticalAlign,
        wrapText: defaultXf.wrapText,
        applyProtection: true,
        hasProtection: true,
        locked: false,
        hidden: true,
        quotePrefix: true,
      });
    } finally {
      mounted.dispose();
      loaded.dispose();
    }
  });
});
