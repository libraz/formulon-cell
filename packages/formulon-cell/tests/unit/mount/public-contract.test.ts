import type { PrinterProfile } from '@libraz/formulon-cell';
import { mount, Workbook } from '@libraz/formulon-cell';
import type { Extension as PublicExtension } from '@libraz/formulon-cell/extensions';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { workbookOf } from '../../../src/api/workbook.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import type { Extension } from '../../../src/extensions/types.js';
import { spreadsheetInternals } from '../../../src/mount/internals.js';
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
    const seed = vi.fn((wb: Workbook) => {
      wb.sheet(0).setNumber(0, 0, 123);
    });
    const instance = await mount(host, { seed });

    try {
      expect(seed).toHaveBeenCalledTimes(1);
      expect(instance.workbook.sheet(0).getValue(0, 0)).toBe(123);
      expect(
        spreadsheetInternals(instance).store.getState().data.cells.get('0:0:0')?.value,
      ).toEqual({
        kind: 'number',
        value: 123,
      });
      expect(instance.history.canUndo).toBe(false);
      instance.history.undo();
      expect(instance.workbook.sheet(0).getValue(0, 0)).toBe(123);
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

  it('normalizes screen clipping hook results before exposing them to UI features', async () => {
    const captureScreenClip = vi
      .fn<() => string | { src: string; alt?: string } | null | undefined>()
      .mockReturnValueOnce('data:image/png;base64,one')
      .mockReturnValueOnce({ src: 'blob:two', alt: 'Selected report area' })
      .mockReturnValueOnce('')
      .mockReturnValueOnce(undefined);
    installDomStubs();
    const { host, cleanup } = createHostElement();
    const instance = await mount(host, { host: { captureScreenClip } });

    try {
      await expect(instance.captureScreenClip()).resolves.toEqual({
        src: 'data:image/png;base64,one',
      });
      await expect(instance.captureScreenClip()).resolves.toEqual({
        src: 'blob:two',
        alt: 'Selected report area',
      });
      await expect(instance.captureScreenClip()).resolves.toBeNull();
      await expect(instance.captureScreenClip()).resolves.toBeNull();
    } finally {
      instance.dispose();
      cleanup();
    }
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
      .fn<() => readonly [typeof nextProfile] | null>()
      .mockReturnValueOnce(null)
      .mockReturnValueOnce(null)
      .mockReturnValueOnce([nextProfile]);
    installDomStubs();
    const { host, cleanup } = createHostElement();
    const instance = await mount(host, {
      host: { printerProfiles: [firstProfile], refreshPrinterProfiles },
    });

    try {
      const kept = await instance.refreshPrinterProfiles();
      expect(kept).toEqual([firstProfile]);
      // The result is a copy: emptying it leaves the list in effect intact.
      (kept as PrinterProfile[]).length = 0;
      await expect(instance.refreshPrinterProfiles()).resolves.toEqual([firstProfile]);
      await expect(instance.refreshPrinterProfiles()).resolves.toEqual([nextProfile]);
    } finally {
      instance.dispose();
      cleanup();
    }
  });

  it('resolves a printer refresh with null when no list was ever supplied', async () => {
    installDomStubs();
    const { host, cleanup } = createHostElement();
    const instance = await mount(host);

    try {
      await expect(instance.refreshPrinterProfiles()).resolves.toBeNull();
    } finally {
      instance.dispose();
      cleanup();
    }
  });

  it('rebases internal instance state when the workbook is replaced', async () => {
    const rebindWorkbook = vi.fn();
    const extension: Extension = {
      id: 'rebind-probe',
      setup(ctx) {
        const stop = ctx.spreadsheet.on('workbook', ({ workbook }) => rebindWorkbook(workbook));
        return { dispose: stop };
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
    expect(rebindWorkbook).toHaveBeenCalledWith(workbookOf(next));
    expect(onWorkbookChange).toHaveBeenCalledTimes(1);
    expect(onWorkbookChange).toHaveBeenCalledWith({ workbook: next });
  });

  it('rebases public instance state when update replaces the workbook', async () => {
    installDomStubs();
    const { host, cleanup } = createHostElement();
    const rebindWorkbook = vi.fn();
    const extension: PublicExtension = {
      id: 'rebind-probe',
      setup(ctx) {
        const stop = ctx.spreadsheet.on('workbook', ({ workbook }) => rebindWorkbook(workbook));
        return { dispose: stop };
      },
    };
    const instance = await mount(host, { extensions: [extension] });
    instance.workbook.addSheet('Second');
    instance.selection.set('Second!A1');
    const next = await Workbook.createDefault();
    next.sheet(0).setNumber(1, 1, 42);
    const onWorkbook = vi.fn();
    const swapped = new Promise<void>((resolve) =>
      instance.on('workbook', (event) => {
        onWorkbook(event);
        resolve();
      }),
    );

    try {
      instance.update({ workbook: next });
      await swapped;

      expect(instance.workbook).toBe(next);
      expect(instance.selection.get().sheet).toBe(0);
      expect(
        spreadsheetInternals(instance).store.getState().data.cells.get('0:1:1')?.value,
      ).toEqual({
        kind: 'number',
        value: 42,
      });
      expect(rebindWorkbook).toHaveBeenCalledWith(next);
      expect(onWorkbook).toHaveBeenCalledTimes(1);
      expect(onWorkbook).toHaveBeenCalledWith({ workbook: next });
    } finally {
      instance.dispose();
      next.dispose();
      cleanup();
    }
  });
});
