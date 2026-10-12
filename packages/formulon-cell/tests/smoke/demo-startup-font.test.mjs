import { describe, expect, it, vi } from 'vitest';
import {
  createInitialDemoWorkbook,
  openDemoWorkbookFile,
} from '../../../../apps/demo-shared/index.ts';
import { WorkbookHandle } from '../../src/engine/workbook-handle.ts';
import { Spreadsheet } from '../../src/mount.ts';
import { createHostElement, installDomStubs, uninstallDomStubs } from '../test-utils/dom.ts';

const columnWidths = (workbook) =>
  workbook
    .getColumnLayouts(0)
    .flatMap(({ first, last, width, hasWidth }) =>
      hasWidth === false
        ? []
        : Array.from({ length: last - first + 1 }, (_, offset) => [first + offset, width]),
    );

describe('demo startup workbook font and seeded layout', () => {
  it('releases a loaded handle when the instance rejects file replacement', async () => {
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const load = vi.spyOn(WorkbookHandle, 'loadBytes').mockResolvedValueOnce(workbook);
    const original = new Error('instance disposed');
    try {
      await expect(
        openDemoWorkbookFile(
          {
            setWorkbook: async () => {
              throw original;
            },
          },
          { name: 'late.xlsx', arrayBuffer: async () => new ArrayBuffer(0) },
        ),
      ).rejects.toBe(original);
      expect(() => workbook.getValue({ sheet: 0, row: 0, col: 0 })).toThrow('disposed');
    } finally {
      load.mockRestore();
      workbook.dispose();
    }
  });
  it.each([
    ['ja', '游ゴシック'],
    ['en', 'Calibri'],
  ])('uses the locale Normal font and seeded column widths for %s', async (locale, fontName) => {
    const workbook = await createInitialDemoWorkbook(`?locale=${locale}`);
    try {
      expect(workbook.isStub).toBe(false);
      expect(workbook.workbookDisplayFont).toEqual({ name: fontName, size: 11 });
      const widths = columnWidths(workbook);
      expect(widths.map(([col]) => col)).toEqual([0, 1, 2, 3]);
      expect(widths[0]?.[1]).toBeCloseTo(112, 6);
      expect(widths[1]?.[1]).toBeCloseTo(80, 6);
      expect(widths[2]?.[1]).toBeCloseTo(104, 6);
      expect(widths[3]?.[1]).toBeCloseTo(128, 6);
    } finally {
      workbook.dispose();
    }
  });

  it('keeps the empty fixture free of seeded values and column overrides', async () => {
    const workbook = await createInitialDemoWorkbook('?locale=ja&fixture=empty');
    try {
      expect(workbook.isStub).toBe(false);
      expect(workbook.workbookDisplayFont).toEqual({ name: '游ゴシック', size: 11 });
      expect(workbook.getColumnLayouts(0)).toEqual([]);
      expect([...workbook.cells(0)]).toEqual([]);
    } finally {
      workbook.dispose();
    }
  });

  it('hands an opened workbook to a live mount and releases both owners on teardown', async () => {
    installDomStubs();
    const source = await WorkbookHandle.createDefault();
    const bytes = source.save();
    source.dispose();

    const original = await WorkbookHandle.createDefault();
    const originalDispose = vi.spyOn(original, 'dispose');
    const { host, cleanup } = createHostElement();
    const instance = await Spreadsheet.mount(host, { workbook: original });
    let adopted;
    let adoptedDispose;
    try {
      const name = await openDemoWorkbookFile(
        instance,
        new File([bytes], 'opened.xlsx', {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }),
      );
      adopted = instance.workbook;
      adoptedDispose = vi.spyOn(adopted, 'dispose');

      expect(name).toBe('opened');
      expect(adopted).not.toBe(original);
      expect(adopted.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });

      // The demo owns the borrowed original after a successful handoff.
      original.dispose();
      expect(originalDispose).toHaveBeenCalledTimes(1);

      // Core owns the workbook adopted by setWorkbook and releases it when
      // the live wrapper instance unmounts.
      instance.dispose();
      expect(adoptedDispose).toHaveBeenCalledTimes(1);
    } finally {
      try {
        instance.dispose();
      } catch {}
      if (!originalDispose.mock.calls.length) original.dispose();
      cleanup();
      uninstallDomStubs();
    }
  });
});
