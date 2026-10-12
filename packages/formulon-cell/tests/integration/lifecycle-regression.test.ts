import { afterEach, describe, expect, it, vi } from 'vitest';

import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { dictionaries } from '../../src/i18n/strings.js';
import { Spreadsheet } from '../../src/mount.js';
import {
  createHostElement,
  installDomStubs,
  mountStubSheet,
  uninstallDomStubs,
} from '../test-utils/index.js';

describe('integration: workbook and extension lifecycle', () => {
  it('restores the older live mount history when the newest shared mount disposes', async () => {
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const first = await mountStubSheet({ workbook });
    const second = await mountStubSheet({ workbook });
    try {
      second.instance.dispose();
      workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 42);
      expect(first.instance.undo()).toBe(true);
      expect(workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
    } finally {
      first.dispose();
      second.dispose();
      workbook.dispose();
    }
  });
  it('shows a localized bounded-print error without creating a print iframe', async () => {
    const sheet = await mountStubSheet({ locale: 'ja' });
    try {
      sheet.workbook.setNumber({ sheet: 0, row: 1000, col: 1000 }, 1);
      expect(() => sheet.instance.print()).not.toThrow();
      expect(document.querySelector('[role="alertdialog"] .fc-tb__dlg__message')?.textContent).toBe(
        dictionaries.ja.pageSetup.printRangeTooLarge,
      );
      expect(sheet.host.querySelector('iframe[data-fc-print-mode]')).toBeNull();
      document
        .querySelector<HTMLButtonElement>('[role="alertdialog"] .fc-fmtdlg__footer button')
        ?.click();
    } finally {
      sheet.dispose();
      sheet.workbook.dispose();
    }
  });
  it('rejects late workbook replacement after disposal without adopting the incoming handle', async () => {
    const sheet = await mountStubSheet();
    const next = await WorkbookHandle.createDefault({ preferStub: true });
    try {
      sheet.instance.dispose();
      await expect(sheet.instance.setWorkbook(next)).rejects.toThrow('disposed');
      expect(next.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'blank' });
    } finally {
      sheet.dispose();
      next.dispose();
    }
  });

  it('preserves the seed failure even when owned-workbook cleanup throws', async () => {
    installDomStubs();
    const owned = await WorkbookHandle.createDefault({ preferStub: true });
    const createDefault = vi.spyOn(WorkbookHandle, 'createDefault').mockResolvedValueOnce(owned);
    const dispose = vi.spyOn(owned, 'dispose').mockImplementationOnce(() => {
      throw new Error('cleanup failed');
    });
    const log = vi.spyOn(console, 'error').mockImplementation(() => {});
    const { host, cleanup } = createHostElement();
    const seedError = new Error('original seed failed');
    try {
      await expect(
        Spreadsheet.mount(host, {
          seed: () => {
            throw seedError;
          },
        }),
      ).rejects.toBe(seedError);
      expect(dispose).toHaveBeenCalledTimes(1);
    } finally {
      createDefault.mockRestore();
      dispose.mockRestore();
      log.mockRestore();
      owned.dispose();
      cleanup();
    }
  });

  it('cleans a fully-built mount when extension setup throws', async () => {
    installDomStubs();
    const owned = await WorkbookHandle.createDefault({ preferStub: true });
    const createDefault = vi.spyOn(WorkbookHandle, 'createDefault').mockResolvedValueOnce(owned);
    const dispose = vi.spyOn(owned, 'dispose');
    const setupError = new Error('extension setup failed');
    const onError = vi.fn();
    const { host, cleanup } = createHostElement();

    try {
      await expect(
        Spreadsheet.mount(host, {
          renderError: false,
          onError,
          extensions: [
            {
              id: 'throws-during-setup',
              setup() {
                throw setupError;
              },
            },
          ],
        }),
      ).rejects.toBe(setupError);
      expect(onError).toHaveBeenCalledTimes(1);
      expect(onError).toHaveBeenCalledWith(setupError);
      expect(dispose).toHaveBeenCalledTimes(1);
      expect(host.children.length).toBe(0);
      expect(host.classList.contains('fc-host')).toBe(false);
    } finally {
      createDefault.mockRestore();
      dispose.mockRestore();
      owned.dispose();
      cleanup();
    }
  });

  it('refreshes the feature view after a throwing setExtensions disposer', async () => {
    const sheet = await mountStubSheet({
      extensions: [
        {
          id: 'old',
          setup: () => ({
            dispose() {
              throw new Error('old cleanup failed');
            },
          }),
        },
      ],
    });

    try {
      expect(() =>
        sheet.instance.setExtensions([
          {
            id: 'next',
            setup: () => ({ dispose() {} }),
          },
        ]),
      ).toThrow('old cleanup failed');
      expect(sheet.instance.features.old).toBeUndefined();
      expect(sheet.instance.features.next).toBeUndefined();

      sheet.instance.setExtensions([{ id: 'next', setup: () => ({ dispose() {} }) }]);
      expect(sheet.instance.features.next).toBeDefined();
    } finally {
      sheet.dispose();
    }
  });
  afterEach(() => {
    uninstallDomStubs();
  });

  it('keeps a newer shared-workbook history attached when an older mount disposes', async () => {
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const first = await mountStubSheet({ workbook });
    const second = await mountStubSheet({ workbook });

    try {
      first.instance.dispose();
      workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 42);

      expect(second.instance.undo()).toBe(true);
      expect(second.instance.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
        kind: 'blank',
      });
    } finally {
      first.dispose();
      second.dispose();
      workbook.dispose();
    }
  });

  it('runs every extension disposer and releases the host when one throws', async () => {
    const calls: string[] = [];
    const firstError = new Error('first disposer failed');
    const sheet = await mountStubSheet({
      extensions: [
        {
          id: 'first',
          setup: () => ({
            dispose() {
              calls.push('first');
              throw firstError;
            },
          }),
        },
        {
          id: 'second',
          setup: () => ({
            dispose() {
              calls.push('second');
            },
          }),
        },
      ],
    });

    try {
      expect(() => sheet.instance.dispose()).toThrow(firstError);
      expect(calls).toEqual(['first', 'second']);
      expect(sheet.host.children.length).toBe(0);
      expect(sheet.host.classList.contains('fc-host')).toBe(false);
    } finally {
      sheet.dispose();
    }
  });

  it('disposes an owned workbook when seed rejects before mount returns', async () => {
    installDomStubs();
    const owned = await WorkbookHandle.createDefault({ preferStub: true });
    const createDefault = vi.spyOn(WorkbookHandle, 'createDefault').mockResolvedValueOnce(owned);
    const { host, cleanup } = createHostElement();
    const seedError = new Error('seed failed');

    try {
      await expect(
        Spreadsheet.mount(host, {
          seed: () => {
            throw seedError;
          },
        }),
      ).rejects.toBe(seedError);
      expect(() => owned.getValue({ sheet: 0, row: 0, col: 0 })).toThrow('disposed');
    } finally {
      createDefault.mockRestore();
      cleanup();
    }
  });
});
