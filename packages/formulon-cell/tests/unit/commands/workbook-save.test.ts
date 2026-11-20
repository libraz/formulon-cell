import { AnchorEditAs, AnchorKind, FormulonError, StatusCode } from '@libraz/formulon';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { SpreadsheetInstance } from '../../../src/api/spreadsheet/instance.js';
import { workbookOf } from '../../../src/api/workbook.js';
import {
  WORKBOOK_SAVE_FORMATS,
  writeWorkbook,
  xlsbUnwritableOf,
} from '../../../src/commands/workbook-save.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { defaultStrings } from '../../../src/i18n/strings.js';
import {
  chooseWorkbookSaveFormat,
  saveWithPrompts,
  workbookSaveFormats,
} from '../../../src/toolbar/dialogs/workbook-save.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

/** A note on the third sheet, which `.xlsb` writes like any other. */
const noteWorkbook = async (): Promise<WorkbookHandle> => {
  const wb = await WorkbookHandle.createDefault();
  wb.addSheet('Sheet2');
  wb.addSheet('Sheet3');
  wb.setText({ sheet: 2, row: 0, col: 0 }, 'x');
  wb.setCommentEntry(2, 0, 0, 'Author', 'A note');
  return wb;
};

/** A shape plus an engine deferred-feature count, which makes an `.xlsb` save lossy. The shape
 *  is written; the count is a unit no category accounts for. */
const insertBox = (wb: WorkbookHandle, deferredFeatureCount: number): void => {
  wb.insertShape(0, {
    name: 'Box',
    presetGeometry: 'rect',
    anchor: {
      kind: AnchorKind.TwoCell,
      editAs: AnchorEditAs.TwoCell,
      fromRow: 1,
      fromCol: 1,
      toRow: 3,
      toCol: 3,
    },
    fill: { color: '#336699' },
  });
  const save = wb.saveWithDiagnostics.bind(wb);
  vi.spyOn(wb, 'saveWithDiagnostics').mockImplementation((format) => {
    const written = save(format);
    return { ...written, diagnostics: { ...written.diagnostics, deferredFeatureCount } };
  });
};

const lossyWorkbook = async (): Promise<WorkbookHandle> => {
  const wb = await WorkbookHandle.createDefault();
  insertBox(wb, 1);
  return wb;
};

/** The parts of a mounted spreadsheet that `saveWithPrompts` reads. */
const spreadsheetOf = (wb: WorkbookHandle): SpreadsheetInstance =>
  ({
    workbook: workbookOf(wb),
    i18n: { strings: defaultStrings },
    host: document.body,
  }) as unknown as SpreadsheetInstance;

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

  it('reports no loss for notes on any sheet', async () => {
    const wb = await noteWorkbook();
    try {
      expect(writeWorkbook(wb, 'xlsb').losses).toEqual([]);
      expect(writeWorkbook(wb, 'xlsx').losses).toEqual([]);
    } finally {
      wb.dispose();
    }
  });

  it('lists drawings as possibly lost and reports deferred settings it cannot itemize', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      insertBox(wb, 1);
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
    const wb = await WorkbookHandle.createDefault();
    try {
      insertBox(wb, 2);
      const pending = saveWithPrompts(spreadsheetOf(wb), { format: 'xlsb' });
      await flush();
      const text = dialog()?.textContent ?? '';
      expect(text).toContain(defaultStrings.workbookSave.drawings);
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
    const wb = await lossyWorkbook();
    try {
      const pending = saveWithPrompts(spreadsheetOf(wb), { format: 'xlsb' });
      await flush();
      expect(dialog()?.textContent).toContain(defaultStrings.workbookSave.otherSettings);
      clickButton(defaultStrings.workbookSave.cancel);
      expect(await pending).toBeNull();
    } finally {
      wb.dispose();
    }
  });

  it('delivers the bytes once the warning is accepted', async () => {
    const wb = await lossyWorkbook();
    try {
      const pending = saveWithPrompts(spreadsheetOf(wb), { format: 'xlsb' });
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
      const outcome = await saveWithPrompts(spreadsheetOf(wb), { format: 'xlsx' });
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

  it('asks for the file type itself when no format is given', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const pending = saveWithPrompts(spreadsheetOf(wb));
      await flush();
      const select = document.querySelector<HTMLSelectElement>('[data-fc-save-format]');
      if (select) select.value = 'xlsb';
      clickButton(defaultStrings.workbookSave.save);
      expect((await pending)?.format).toBe('xlsb');
    } finally {
      wb.dispose();
    }
  });

  it('resolves null when the file type choice is cancelled', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const pending = saveWithPrompts(spreadsheetOf(wb));
      await flush();
      clickButton(defaultStrings.workbookSave.cancel);
      expect(await pending).toBeNull();
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

describe.skipIf(!canLoadWasm())('workbook save refused by the .xlsb writer', () => {
  /** A protected range with a security descriptor on Sheet2, whose SDDL is kept as given, which `.xlsb` cannot hold. */
  const refusedWorkbook = async (): Promise<WorkbookHandle> => {
    const wb = await WorkbookHandle.createDefault();
    wb.addSheet('Sheet2');
    wb.addProtectedRange(1, {
      name: 'Edit',
      ranges: 'A1:B2',
      securityDescriptor: 'D:(A;;GA;;;S-1-5-21-1-2-3-500)',
    });
    return wb;
  };

  it('reads the feature and sheet from the refusal and nothing else', () => {
    const refusal = (context: string) =>
      new FormulonError(StatusCode.IoFeatureNotWritable, 'not writable', context);
    expect(xlsbUnwritableOf(refusal('context=write_xlsb kind=sparklines sheet_index=3'))).toEqual({
      id: 'sparklines',
      sheet: 3,
    });
    expect(
      xlsbUnwritableOf(refusal('context=write_xlsb kind=protected_range_rects sheet_index=0')),
    ).toEqual({ id: 'protectedRanges', sheet: 0 });
    expect(xlsbUnwritableOf(refusal('context=write_xlsb kind=data_bar sheet_index=1'))).toEqual({
      id: 'dataBars',
      sheet: 1,
    });
    expect(xlsbUnwritableOf(refusal('context=write_xlsb kind=x14_data_bar'))).toEqual({
      id: 'unknownFeature',
      sheet: null,
    });
    expect(xlsbUnwritableOf(new FormulonError(StatusCode.IoWriteFailed, 'io', ''))).toBeNull();
    expect(xlsbUnwritableOf(new Error('x'))).toBeNull();
  });

  it('names the feature and its sheet, and saves as .xlsx when asked', async () => {
    const wb = await refusedWorkbook();
    try {
      const pending = saveWithPrompts(spreadsheetOf(wb), { format: 'xlsb' });
      await flush();
      const t = defaultStrings.workbookSave;
      const text = dialog()?.textContent ?? '';
      expect(text).toContain(t.unwritableMessage);
      expect(text).toContain(`${t.protectedRanges} (Sheet2)`);
      clickButton(t.saveAsFormat.replace('{format}', t.formatXlsx));
      const outcome = await pending;
      expect(outcome?.format).toBe('xlsx');
      expect(outcome?.bytes.length).toBeGreaterThan(0);
    } finally {
      wb.dispose();
    }
  });

  it('returns null when the notice is cancelled', async () => {
    const wb = await refusedWorkbook();
    try {
      const pending = saveWithPrompts(spreadsheetOf(wb), { format: 'xlsb' });
      await flush();
      clickButton(defaultStrings.workbookSave.cancel);
      expect(await pending).toBeNull();
    } finally {
      wb.dispose();
    }
  });
});
