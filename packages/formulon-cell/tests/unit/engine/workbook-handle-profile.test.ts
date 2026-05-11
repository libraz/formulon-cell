import { describe, expect, it } from 'vitest';
import type { FormulonModule, Workbook } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const makeHandle = (supportsProfile: boolean): { wb: WorkbookHandle; setCalls: string[] } => {
  const setCalls: string[] = [];
  const raw = supportsProfile
    ? {
        excelProfileId: () => 'win-365-ja_JP',
        setExcelProfileId: (profileId: string) => {
          setCalls.push(profileId);
          return { ok: true, code: 0, message: '' };
        },
      }
    : {};
  const module = { versionString: () => 'test' } as unknown as FormulonModule;
  const Ctor = WorkbookHandle as unknown as new (
    module: FormulonModule,
    wb: Workbook,
  ) => WorkbookHandle;
  return { wb: new Ctor(module, raw as unknown as Workbook), setCalls };
};

describe('WorkbookHandle spreadsheet profile', () => {
  it('reads and writes the engine formula-behaviour profile when available', () => {
    const { wb, setCalls } = makeHandle(true);

    expect(wb.capabilities.spreadsheetProfile).toBe(true);
    expect(wb.spreadsheetProfileId()).toBe('windows-ja_JP');
    expect(wb.setSpreadsheetProfileId('mac-ja_JP')).toBe(true);
    expect(setCalls).toEqual(['mac-365-ja_JP']);
  });

  it('is a null/false no-op when the engine package is older', () => {
    const { wb, setCalls } = makeHandle(false);

    expect(wb.capabilities.spreadsheetProfile).toBe(false);
    expect(wb.spreadsheetProfileId()).toBeNull();
    expect(wb.setSpreadsheetProfileId('mac-ja_JP')).toBe(false);
    expect(setCalls).toEqual([]);
  });
});
