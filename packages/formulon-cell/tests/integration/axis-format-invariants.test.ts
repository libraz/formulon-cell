import { describe, expect, it } from 'vitest';
import { applySelectionFormatAction } from '../../src/commands/format.js';
import { History } from '../../src/commands/history.js';
import { recordFormatChange } from '../../src/commands/slice-history.js';
import { MAX_COL, MAX_ROW } from '../../src/engine/address.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { mutators } from '../../src/store/store.js';
import { mountStubSheet } from '../test-utils/mount.js';

function nativeFormat(workbook: WorkbookHandle, row: number, col: number) {
  const effective = workbook.getEffectiveStyle(0, row, col);
  const xf = effective && workbook.getCellXf(effective.xfIndex);
  if (!xf) throw new Error(`Missing native XF at ${row}:${col}`);
  const font = workbook.getFontRecord(xf.fontIndex);
  const fill = workbook.getFillRecord(xf.fillIndex);
  return {
    size: font?.size,
    name: font?.name,
    bold: font?.bold,
    italic: font?.italic,
    underline: font?.underline,
    strike: font?.strike,
    color: font?.colorArgb,
    fill: fill?.fgArgb,
    horizontal: xf.horizontalAlign,
    vertical: xf.verticalAlign,
    wrap: xf.wrapText,
    rotation: xf.textRotation,
    shrink: xf.shrinkToFit,
    number: workbook.getNumFmtCode(xf.numFmtId),
    quotePrefix: xf.quotePrefix,
  };
}

describe('native axis format invariants', () => {
  for (const axis of ['column', 'rows'] as const) {
    it(`keeps scalar formatting, future input and XLSX consistent for ${axis}`, async () => {
      const workbook = await WorkbookHandle.createDefault({ locale: 'ja' });
      const mounted = await mountStubSheet({ workbook, locale: 'ja' });
      const history = new History();
      const far = axis === 'column' ? [MAX_ROW, 1] : [9, MAX_COL];
      const expected = {
        size: 18,
        name: 'Arial',
        bold: true,
        italic: true,
        underline: 2,
        strike: true,
        color: 0xffff0000,
        fill: 0xff00ff00,
        horizontal: 2,
        vertical: 0,
        wrap: true,
        rotation: 30,
        shrink: true,
        number: '0.00',
      };
      try {
        expect(workbook.isStub).toBe(false);
        if (axis === 'column') mutators.selectCol(mounted.instance.store, 1);
        else mutators.selectRows(mounted.instance.store, 2, 9);
        recordFormatChange(history, mounted.instance.store, () => {
          expect(
            applySelectionFormatAction(mounted.instance.store.getState(), mounted.instance.store, {
              patch: {
                fontSize: 24,
                fontFamily: 'Arial',
                bold: true,
                italic: true,
                underline: 'double',
                strike: true,
                color: '#ff0000',
                fill: '#00ff00',
                align: 'center',
                vAlign: 'top',
                wrap: true,
                rotation: 30,
                shrinkToFit: true,
                numFmt: { kind: 'fixed', decimals: 2 },
              },
            }),
          ).toBe(true);
        });
        expect(nativeFormat(workbook, far[0] as number, far[1] as number)).toMatchObject(expected);
        expect(nativeFormat(workbook, 0, 0).size).toBe(11);
        expect(
          mounted.instance.applyChanges([{ addr: { sheet: 0, row: 9, col: 1 }, input: '12.5' }])
            .status,
        ).toBe('applied');
        expect(workbook.getValue({ sheet: 0, row: 9, col: 1 })).toEqual({
          kind: 'number',
          value: 12.5,
        });
        expect(nativeFormat(workbook, 9, 1)).toMatchObject(expected);
        const reloaded = await WorkbookHandle.loadBytes(workbook.save());
        try {
          expect(nativeFormat(reloaded, 9, 1)).toMatchObject(expected);
          expect(nativeFormat(reloaded, far[0] as number, far[1] as number)).toMatchObject(
            expected,
          );
          expect(nativeFormat(reloaded, 0, 0).size).toBe(11);
        } finally {
          reloaded.dispose();
        }
        expect(history.undo()).toBe(true);
        expect(nativeFormat(workbook, far[0] as number, far[1] as number).size).toBe(11);
        expect(workbook.getValue({ sheet: 0, row: 9, col: 1 })).toEqual({
          kind: 'number',
          value: 12.5,
        });
        expect(history.redo()).toBe(true);
        expect(nativeFormat(workbook, far[0] as number, far[1] as number)).toMatchObject(expected);
      } finally {
        mounted.dispose();
        workbook.dispose();
      }
    });
  }

  it.each([false, true])(
    'includes an opaque native direct XF (hasValue=%s) in a column edit',
    async (hasValue) => {
      const workbook = await WorkbookHandle.createDefault({ locale: 'ja' });
      const normal = workbook.getCellXf(0);
      if (!normal) throw new Error('Missing native Normal XF');
      const direct = workbook.addXfRecord({ ...normal, quotePrefix: true });
      expect(direct).toBeGreaterThan(0);
      if (hasValue) workbook.setText({ sheet: 0, row: 1, col: 1 }, 'metadata');
      expect(workbook.setCellXfIndex(0, 1, 1, direct)).toBe(true);
      expect([...workbook.physicalCells(0)].map((cell) => cell.addr)).toContainEqual({
        sheet: 0,
        row: 1,
        col: 1,
      });
      const mounted = await mountStubSheet({ workbook, locale: 'ja' });
      try {
        expect(mounted.instance.store.getState().format.formats.has('0:1:1')).toBe(false);
        expect(nativeFormat(workbook, 1, 1)).toMatchObject({ size: 11, quotePrefix: true });
        mutators.selectCol(mounted.instance.store, 1);
        expect(
          applySelectionFormatAction(mounted.instance.store.getState(), mounted.instance.store, {
            patch: { fontSize: 24 },
          }),
        ).toBe(true);
        expect(nativeFormat(workbook, 1, 1)).toMatchObject({ size: 18, quotePrefix: true });
        expect(nativeFormat(workbook, MAX_ROW, 1).size).toBe(18);
        expect(workbook.getValue({ sheet: 0, row: 1, col: 1 })).toEqual(
          hasValue ? { kind: 'text', value: 'metadata' } : { kind: 'blank' },
        );
      } finally {
        mounted.dispose();
        workbook.dispose();
      }
    },
  );

  it('preserves a later single-cell edit when another column intersects an earlier row', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'ja' });
    workbook.setText({ sheet: 0, row: 2, col: 1 }, 'keep');
    const mounted = await mountStubSheet({ workbook, locale: 'ja' });
    const patch = (fontSize: number) =>
      expect(
        applySelectionFormatAction(mounted.instance.store.getState(), mounted.instance.store, {
          patch: { fontSize },
        }),
      ).toBe(true);
    try {
      mutators.selectRow(mounted.instance.store, 2);
      patch(24);
      mutators.selectCol(mounted.instance.store, 1);
      patch(32);
      expect(nativeFormat(workbook, 2, 1).size).toBe(24);
      mutators.setActive(mounted.instance.store, { sheet: 0, row: 2, col: 1 });
      patch(40);
      expect(nativeFormat(workbook, 2, 1).size).toBe(30);
      mutators.selectCol(mounted.instance.store, 3);
      patch(24);
      expect(nativeFormat(workbook, 2, 3).size).toBe(18);
      expect(nativeFormat(workbook, 2, 1).size).toBe(30);
      expect(workbook.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({
        kind: 'text',
        value: 'keep',
      });
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('restores the inherited appearance on undo after a blank intersection receives a value', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'ja' });
    const normal = workbook.getCellXf(0);
    const font = normal && workbook.getFontRecord(normal.fontIndex);
    if (!normal || !font) throw new Error('Missing native Normal font');
    const fontIndex = workbook.addFontRecord({ ...font, size: 18 });
    const rowXf = workbook.addXfRecord({ ...normal, fontIndex, applyFont: true });
    expect(workbook.setRowStyle(0, 2, rowXf)).toBe(true);
    const mounted = await mountStubSheet({ workbook, locale: 'ja' });
    const history = new History();
    try {
      expect(nativeFormat(workbook, 2, 1).size).toBe(18);
      mutators.selectCol(mounted.instance.store, 1);
      recordFormatChange(history, mounted.instance.store, () => {
        expect(
          applySelectionFormatAction(mounted.instance.store.getState(), mounted.instance.store, {
            patch: { fontSize: 32 },
          }),
        ).toBe(true);
      });
      expect(nativeFormat(workbook, 2, 1).size).toBe(24);
      expect(
        mounted.instance.applyChanges([{ addr: { sheet: 0, row: 2, col: 1 }, input: '12.5' }])
          .status,
      ).toBe('applied');
      expect(history.undo()).toBe(true);
      expect(nativeFormat(workbook, 2, 1).size).toBe(18);
      expect(workbook.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({
        kind: 'number',
        value: 12.5,
      });
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });
});

describe('implicit column style preservation', () => {
  for (const explicitNormal of [false, true]) {
    it(`keeps column explicitness through row edits and column undo (${explicitNormal})`, async () => {
      const workbook = await WorkbookHandle.createDefault({ locale: 'ja' });
      if (explicitNormal) expect(workbook.setColumnStyle(0, 2, 2, 0)).toBe(true);
      const mounted = await mountStubSheet({ workbook, locale: 'ja' });
      const history = new History();
      const store = mounted.instance.store;
      const before = workbook.getSheetFontAxisSnapshot(0)?.columns;
      try {
        expect(before).toEqual(explicitNormal ? [{ first: 2, last: 2, styleXf: 0 }] : []);
        mutators.selectRow(store, 2);
        recordFormatChange(history, store, () => {
          expect(
            applySelectionFormatAction(store.getState(), store, {
              patch: { fontSize: 24 },
            }),
          ).toBe(true);
        });
        expect(workbook.getSheetFontAxisSnapshot(0)?.columns).toEqual(before);
        mutators.selectCol(store, 3);
        recordFormatChange(history, store, () => {
          expect(
            applySelectionFormatAction(store.getState(), store, {
              patch: { color: '#ff0000' },
            }),
          ).toBe(true);
        });
        expect(history.undo()).toBe(true);
        expect(workbook.getSheetFontAxisSnapshot(0)?.columns).toEqual(before);
        expect(workbook.getEffectiveStyle(0, 3, 0)?.source).toBe(3);
        expect(workbook.getEffectiveStyle(0, 3, 2)?.source).toBe(explicitNormal ? 2 : 3);
        expect(history.undo()).toBe(true);
        expect(workbook.getSheetFontAxisSnapshot(0)?.columns).toEqual(before);
        expect(workbook.getEffectiveStyle(0, 2, 4)?.source).toBe(3);
      } finally {
        mounted.dispose();
        workbook.dispose();
      }
    });
  }
});
