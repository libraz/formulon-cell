import { afterEach, describe, expect, it, vi } from 'vitest';
import type { CellValue } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import type { Extension } from '../../../src/extensions/types.js';
import { Spreadsheet } from '../../../src/mount.js';
import { mutators } from '../../../src/store/store.js';
import { createHostElement, installDomStubs, uninstallDomStubs } from '../../test-utils/dom.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

describe('mount/public contract', () => {
  let sheet: MountedStubSheet | undefined;

  afterEach(() => {
    sheet?.dispose();
    sheet = undefined;
    uninstallDomStubs();
  });

  it('runs the demo seed hook for a default-owned workbook without adding undo history', async () => {
    installDomStubs();
    const { host, cleanup } = createHostElement();
    const seed = vi.fn((wb: WorkbookHandle) => {
      wb.setNumber({ sheet: 0, row: 0, col: 0 }, 123);
    });
    const instance = await Spreadsheet.mount(host, { seed });

    try {
      expect(seed).toHaveBeenCalledTimes(1);
      expect(instance.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
        kind: 'number',
        value: 123,
      });
      expect(instance.store.getState().data.cells.get('0:0:0')?.value).toEqual({
        kind: 'number',
        value: 123,
      });
      expect(instance.undo()).toBe(false);
      expect(instance.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
        kind: 'number',
        value: 123,
      });
    } finally {
      instance.dispose();
      cleanup();
    }
  });

  it('does not run the demo seed hook against a caller-owned workbook', async () => {
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    workbook.setText({ sheet: 0, row: 0, col: 0 }, 'caller data');
    const seed = vi.fn((wb: WorkbookHandle) => {
      wb.setText({ sheet: 0, row: 0, col: 0 }, 'seed data');
    });

    sheet = await mountStubSheet({ workbook, seed });

    expect(seed).not.toHaveBeenCalled();
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'caller data',
    });
    expect(sheet.instance.store.getState().data.cells.get('0:0:0')?.value).toEqual({
      kind: 'text',
      value: 'caller data',
    });
  });

  it('registers initial host custom functions case-insensitively with metadata', async () => {
    const double = vi.fn((value: CellValue) => (value.kind === 'number' ? value.value * 2 : null));
    sheet = await mountStubSheet({
      functions: [
        {
          name: 'host_double',
          impl: double,
          meta: {
            description: 'Doubles a number',
            args: ['value'],
            returnType: 'number',
          },
        },
      ],
    });

    expect(sheet.instance.formula.has('HOST_DOUBLE')).toBe(true);
    expect(sheet.instance.formula.has('host_double')).toBe(true);
    expect(sheet.instance.formula.list()).toContain('HOST_DOUBLE');
    expect(sheet.instance.formula.get('HOST_DOUBLE')?.meta).toEqual({
      description: 'Doubles a number',
      args: ['value'],
      returnType: 'number',
    });
    expect(sheet.instance.formula.evaluate('host_double', [{ kind: 'number', value: 21 }])).toEqual(
      {
        kind: 'number',
        value: 42,
      },
    );
    expect(double).toHaveBeenCalledWith({ kind: 'number', value: 21 });
  });

  it('normalizes screen clipping hook results before exposing them to UI features', async () => {
    const captureScreenClip = vi
      .fn<() => string | { src: string; alt?: string } | null | undefined>()
      .mockReturnValueOnce('data:image/png;base64,one')
      .mockReturnValueOnce({ src: 'blob:two', alt: 'Selected report area' })
      .mockReturnValueOnce('')
      .mockReturnValueOnce(undefined);
    sheet = await mountStubSheet({ captureScreenClip });

    await expect(sheet.instance.captureScreenClip()).resolves.toEqual({
      src: 'data:image/png;base64,one',
    });
    await expect(sheet.instance.captureScreenClip()).resolves.toEqual({
      src: 'blob:two',
      alt: 'Selected report area',
    });
    await expect(sheet.instance.captureScreenClip()).resolves.toBeNull();
    await expect(sheet.instance.captureScreenClip()).resolves.toBeNull();
  });

  it('keeps the previous printer profile list when a host refresh has no update', async () => {
    const firstProfile = {
      id: 'office-a4',
      name: 'Office A4',
      paperSize: 'A4',
      orientation: 'portrait',
      printableBounds: { left: 10, top: 12, right: 200, bottom: 285 },
    } as const;
    const nextProfile = {
      id: 'pdf-letter',
      name: 'PDF Letter',
      paperSize: 'letter',
      orientation: 'landscape',
      printableBounds: { left: 0, top: 0, right: 279, bottom: 216 },
    } as const;
    const refreshPrinterProfiles = vi
      .fn<() => readonly [typeof nextProfile] | undefined>()
      .mockReturnValueOnce(undefined)
      .mockReturnValueOnce([nextProfile]);
    sheet = await mountStubSheet({
      printerProfiles: [firstProfile],
      refreshPrinterProfiles,
    });

    await expect(sheet.instance.refreshPrinterProfiles()).resolves.toEqual([firstProfile]);
    await expect(sheet.instance.refreshPrinterProfiles()).resolves.toEqual([nextProfile]);
  });

  it('rebases public instance state when the workbook is replaced', async () => {
    const rebindWorkbook = vi.fn();
    const extension: Extension = {
      id: 'rebind-probe',
      setup() {
        return {
          dispose() {},
          rebindWorkbook,
        };
      },
    };
    sheet = await mountStubSheet({ extensions: [extension] });
    sheet.workbook.addSheet('Second');
    mutators.setSheetIndex(sheet.instance.store, 1);
    const next = await WorkbookHandle.createDefault({ preferStub: true });
    next.setNumber({ sheet: 0, row: 1, col: 1 }, 42);
    const onWorkbookChange = vi.fn();
    sheet.instance.on('workbookChange', onWorkbookChange);

    await sheet.instance.setWorkbook(next);

    expect(sheet.instance.workbook).toBe(next);
    expect(sheet.instance.store.getState().data.sheetIndex).toBe(0);
    expect(sheet.instance.store.getState().data.cells.get('0:1:1')?.value).toEqual({
      kind: 'number',
      value: 42,
    });
    expect(rebindWorkbook).toHaveBeenCalledWith(next);
    expect(onWorkbookChange).toHaveBeenCalledTimes(1);
    expect(onWorkbookChange).toHaveBeenCalledWith({ workbook: next });
  });
});
