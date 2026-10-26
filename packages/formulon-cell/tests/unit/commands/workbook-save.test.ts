import { AnchorEditAs, AnchorKind, ShapePaint } from '@libraz/formulon';
import { afterEach, describe, expect, it } from 'vitest';
import { WORKBOOK_SAVE_FORMATS, writeWorkbook } from '../../../src/commands/workbook-save.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { defaultStrings } from '../../../src/i18n/strings.js';
import {
  chooseWorkbookSaveFormat,
  saveWorkbookWithPrompts,
  workbookSaveFormats,
} from '../../../src/toolbar/dialogs/workbook-save.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

const noteWorkbook = async (): Promise<WorkbookHandle> => {
  const wb = await WorkbookHandle.createDefault();
  wb.setText({ sheet: 0, row: 0, col: 0 }, 'x');
  wb.setCommentEntry(0, 0, 0, 'Author', 'A note');
  return wb;
};

const dialog = (): HTMLElement | null => document.querySelector('[role="alertdialog"]');

const clickButton = (label: string): void => {
  const button = [...document.querySelectorAll<HTMLButtonElement>('button')].find(
    (b) => b.textContent === label,
  );
  if (!button) throw new Error(`no button ${label}`);
  button.click();
};

const flush = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 0));

afterEach(() => {
  for (const el of document.querySelectorAll('[role="alertdialog"], [role="dialog"]')) {
    el.closest('.fc-tb__dlg-overlay')?.remove();
    el.remove();
  }
});

describe.skipIf(!canLoadWasm())('workbook save', () => {
  it('names the file type of each container', () => {
    expect(WORKBOOK_SAVE_FORMATS.xlsx.extension).toBe('.xlsx');
    expect(WORKBOOK_SAVE_FORMATS.xlsb.extension).toBe('.xlsb');
    expect(WORKBOOK_SAVE_FORMATS.xlsb.mimeType).toBe(
      'application/vnd.ms-excel.sheet.binary.macroEnabled.12',
    );
  });

  it('writes .xlsb bytes that load back and reports nothing lost for plain content', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      wb.setNumber({ sheet: 0, row: 0, col: 0 }, 42);
      const outcome = writeWorkbook(wb, 'xlsb');
      expect(outcome.format).toBe('xlsb');
      expect(outcome.losses).toEqual([]);
      const reloaded = await WorkbookHandle.loadBytes(outcome.bytes);
      try {
        expect(reloaded.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
          kind: 'number',
          value: 42,
        });
      } finally {
        reloaded.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('itemizes notes the binary format drops, next to the engine counter', async () => {
    const wb = await noteWorkbook();
    try {
      const outcome = writeWorkbook(wb, 'xlsb');
      expect(outcome.losses).toEqual([
        { id: 'notes', count: 1, certain: true },
        { id: 'deferredFeature', count: 1, certain: true },
      ]);
      expect(writeWorkbook(wb, 'xlsx').losses).toEqual([]);
    } finally {
      wb.dispose();
    }
  });

  it('lists drawings as possibly lost and reports deferred settings it cannot itemize', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      wb.insertShape(0, {
        name: 'Box',
        prst: 'rect',
        anchorKind: AnchorKind.TwoCell,
        editAs: AnchorEditAs.TwoCell,
        fromRow: 1,
        fromCol: 1,
        toRow: 3,
        toCol: 3,
        fillKind: ShapePaint.Solid,
        fillRgb: 0x336699,
      });
      wb.setSheetHeaderFooter(0, { oddHeader: 'Title' });
      const outcome = writeWorkbook(wb, 'xlsb');
      expect(outcome.losses).toEqual(
        expect.arrayContaining([
          { id: 'drawings', count: 1, certain: false },
          { id: 'otherSettings', certain: false },
        ]),
      );
      expect(outcome.losses.some((loss) => loss.id === 'notes')).toBe(false);
    } finally {
      wb.dispose();
    }
  });

  it('separates content that will be lost from content that may be', async () => {
    const wb = await noteWorkbook();
    try {
      wb.setSheetHeaderFooter(0, { oddHeader: 'Title' });
      const pending = saveWorkbookWithPrompts({ wb, strings: defaultStrings, format: 'xlsb' });
      await flush();
      const text = dialog()?.textContent ?? '';
      expect(text).toContain(defaultStrings.workbookSave.notes);
      expect(text).toContain(
        defaultStrings.workbookSave.mayLoseMessage.replace(
          '{format}',
          defaultStrings.workbookSave.formatXlsb,
        ),
      );
      expect(text).toContain(defaultStrings.workbookSave.otherSettings);
      clickButton(defaultStrings.workbookSave.cancel);
      await pending;
    } finally {
      wb.dispose();
    }
  });

  it('asks before delivering a lossy .xlsb and returns null when cancelled', async () => {
    const wb = await noteWorkbook();
    try {
      const pending = saveWorkbookWithPrompts({ wb, strings: defaultStrings, format: 'xlsb' });
      await flush();
      expect(dialog()?.textContent).toContain(defaultStrings.workbookSave.notes);
      clickButton(defaultStrings.workbookSave.cancel);
      expect(await pending).toBeNull();
    } finally {
      wb.dispose();
    }
  });

  it('delivers the bytes once the warning is accepted', async () => {
    const wb = await noteWorkbook();
    try {
      const pending = saveWorkbookWithPrompts({ wb, strings: defaultStrings, format: 'xlsb' });
      await flush();
      clickButton(defaultStrings.workbookSave.saveAnyway);
      const outcome = await pending;
      expect(outcome?.format).toBe('xlsb');
      expect(outcome?.bytes.length).toBeGreaterThan(0);
    } finally {
      wb.dispose();
    }
  });

  it('saves a lossless workbook without any prompt', async () => {
    const wb = await noteWorkbook();
    try {
      const outcome = await saveWorkbookWithPrompts({
        wb,
        strings: defaultStrings,
        format: 'xlsx',
      });
      expect(dialog()).toBeNull();
      expect(outcome?.format).toBe('xlsx');
    } finally {
      wb.dispose();
    }
  });

  it('offers .xlsb in Save As and returns the chosen type', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(workbookSaveFormats(wb)).toEqual(['xlsx', 'xlsb']);
      const pending = chooseWorkbookSaveFormat(wb, defaultStrings);
      await flush();
      const select = document.querySelector<HTMLSelectElement>('[data-fc-save-format]');
      expect([...(select?.options ?? [])].map((o) => o.value)).toEqual(['xlsx', 'xlsb']);
      if (select) select.value = 'xlsb';
      clickButton(defaultStrings.workbookSave.save);
      expect(await pending).toBe('xlsb');
    } finally {
      wb.dispose();
    }
  });

  it('an engine that cannot write .xlsb offers only .xlsx and refuses an .xlsb save', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    try {
      expect(workbookSaveFormats(wb)).toEqual(['xlsx']);
      expect(() => wb.saveWithDiagnostics('xlsb')).toThrow(/xlsb/);
    } finally {
      wb.dispose();
    }
  });
});
