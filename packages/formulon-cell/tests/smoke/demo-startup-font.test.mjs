import { mount } from '@libraz/formulon-cell';
import { describe, expect, it, vi } from 'vitest';
import {
  adoptDemoWorkbook,
  createInitialDemoWorkbook,
  openDemoWorkbookFile,
} from '../../../../apps/demo-shared/index.ts';
import { createHostElement, installDomStubs, uninstallDomStubs } from '../test-utils/dom.ts';

const SEEDED_WIDTHS_PT = [84, 60, 78, 96];

const columnWidthsPt = (workbook, count) =>
  Array.from({ length: count }, (_, col) => workbook.sheet(0).getColumnWidthPt(col));

const xlsxFile = async (name) => {
  const source = await createInitialDemoWorkbook('?locale=en&fixture=empty');
  const bytes = source.save().bytes;
  source.dispose();
  return new File([bytes], name, {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
};

describe('demo startup workbook font and seeded layout', () => {
  it.each([
    ['ja', '游ゴシック'],
    ['en', 'Calibri'],
  ])('uses the locale Normal font and seeded column widths for %s', async (locale, fontName) => {
    const workbook = await createInitialDemoWorkbook(`?locale=${locale}`);
    try {
      expect(workbook.engine).toBe('wasm');
      expect(workbook.getDefaultFont()).toMatchObject({ name: fontName, sizePt: 11 });
      const widths = columnWidthsPt(workbook, 4);
      for (const [col, pt] of SEEDED_WIDTHS_PT.entries()) expect(widths[col]).toBeCloseTo(pt, 0);
    } finally {
      workbook.dispose();
    }
  });

  it('keeps the empty fixture free of seeded values and column overrides', async () => {
    const workbook = await createInitialDemoWorkbook('?locale=ja&fixture=empty');
    try {
      expect(workbook.engine).toBe('wasm');
      expect(workbook.getDefaultFont()).toMatchObject({ name: '游ゴシック', sizePt: 11 });
      const standard = workbook.sheet(0).getColumnWidthPt(20);
      expect(columnWidthsPt(workbook, 4)).toEqual([standard, standard, standard, standard]);
      expect(workbook.sheet(0).listCells()).toEqual([]);
    } finally {
      workbook.dispose();
    }
  });
});

describe('opening a workbook file in the demo', () => {
  it('reads the file into a workbook the caller owns and names it without its extension', async () => {
    const { workbook, name } = await openDemoWorkbookFile(await xlsxFile('opened.xlsx'));
    try {
      expect(name).toBe('opened');
      expect(workbook.disposed).toBe(false);
      expect(workbook.sheet(0).listCells()).toEqual([]);
    } finally {
      workbook.dispose();
    }
  });

  it('releases the opened workbook when the instance refuses it', async () => {
    const { workbook } = await openDemoWorkbookFile(await xlsxFile('refused.xlsx'));
    const refusal = new Error('instance disposed');
    const live = {
      workbook: null,
      on: () => () => {},
      update: () => {
        throw refusal;
      },
    };
    await expect(adoptDemoWorkbook(live, workbook, () => true)).rejects.toBe(refusal);
    expect(workbook.disposed).toBe(true);
  });

  it('releases the opened workbook when the demo goes away before the swap', async () => {
    const { workbook } = await openDemoWorkbookFile(await xlsxFile('late.xlsx'));
    const update = vi.fn();
    const live = { workbook: null, on: () => () => {}, update };
    await expect(adoptDemoWorkbook(live, workbook, () => false)).resolves.toBe(false);
    expect(update).not.toHaveBeenCalled();
    expect(workbook.disposed).toBe(true);
  });

  it('hands the opened workbook to a live mount and leaves it with the caller', async () => {
    installDomStubs();
    const original = await createInitialDemoWorkbook('?locale=en&fixture=empty');
    const { host, cleanup } = createHostElement();
    const instance = await mount(host, { workbook: original });
    const { workbook: opened } = await openDemoWorkbookFile(await xlsxFile('handoff.xlsx'));
    try {
      await expect(adoptDemoWorkbook(instance, opened, () => true)).resolves.toBe(true);
      expect(instance.workbook).toBe(opened);
      expect(opened.disposed).toBe(false);
      // The caller releases the workbook the instance no longer shows.
      original.dispose();
      expect(original.disposed).toBe(true);
    } finally {
      instance.dispose();
      if (!opened.disposed) opened.dispose();
      if (!original.disposed) original.dispose();
      cleanup();
      uninstallDomStubs();
    }
  });
});
