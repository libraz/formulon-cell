import { StatusCode } from '@libraz/formulon';
import { rangeOfRect, rectOfRange } from '../engine/address.js';
import type { CellValue, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';

export interface PivotSourceField {
  name: string;
  index: number;
  /** Whether the field's source column holds a number. */
  numeric: boolean;
}

export interface RefreshPivotCacheOptions {
  cacheId: number;
  source: Range;
}

export interface RefreshPivotTableOptions {
  sheet: number;
  pivotIndex: number;
  source: Range;
}

export interface RefreshPivotTableFromStoredSourceOptions {
  sheet: number;
  pivotIndex: number;
}

export type RefreshPivotCacheResult =
  | { ok: true; cacheId: number }
  | {
      ok: false;
      reason: 'unsupported' | 'invalid-range' | 'invalid-field' | 'invalid-pivot' | 'engine-failed';
      step?: string;
    };

const cellLabel = (value: CellValue, fallback: string): string => {
  if (value.kind === 'text') return value.value.trim() || fallback;
  if (value.kind === 'number') return String(value.value);
  if (value.kind === 'bool') return value.value ? 'TRUE' : 'FALSE';
  return fallback;
};

export const writePivotCacheWorksheetSource = (
  wb: WorkbookHandle,
  cacheId: number,
  source: Range,
): boolean => {
  if (!wb.capabilities.pivotCacheSource) return true;
  return wb.setPivotCacheWorksheetSource(cacheId, {
    range: {
      sheet: source.sheet,
      rect: rectOfRange(source),
    },
  });
};

export const pivotCacheWorksheetSourceRange = (
  wb: WorkbookHandle,
  cacheId: number,
): Range | null => {
  const range = wb.getPivotCacheWorksheetSource(cacheId)?.range;
  if (!range) return null;
  return rangeOfRect(range.sheet, range.rect);
};

/** The fields of a pivot cache, as the engine named and typed them from the source's header row. */
export const pivotCacheSourceFields = (wb: WorkbookHandle, cacheId: number): PivotSourceField[] =>
  wb
    .pivotCacheFields(cacheId)
    .map((field, index) => ({ name: field.name, index, numeric: field.containsNumber }));

/** Whether the source column behind a cache field holds numbers. */
export function pivotCacheFieldIsNumeric(
  wb: WorkbookHandle,
  cacheId: number,
  cacheFieldIndex: number,
): boolean {
  return wb.pivotCacheFields(cacheId)[cacheFieldIndex]?.containsNumber ?? false;
}

/** The distinct items of a cache field, in cache order, as labels. */
export const pivotCacheItemLabels = (
  wb: WorkbookHandle,
  cacheId: number,
  cacheFieldIndex: number,
): string[] =>
  wb.pivotCacheSharedItems(cacheId, cacheFieldIndex).map((v) => cellLabel(v, '(blank)'));

/** A refusal of the engine's Refresh: a source that does not resolve or is too large is the
 *  range's fault, anything else the engine's. */
export const refreshFailure = (
  code: number,
): { ok: false; reason: 'invalid-range' | 'engine-failed'; step: string } => ({
  ok: false,
  reason:
    code === StatusCode.PreconditionFailed || code === StatusCode.SecurityResourceLimit
      ? 'invalid-range'
      : 'engine-failed',
  step: 'refresh',
});

export function refreshPivotCacheFromRange(
  wb: WorkbookHandle,
  opts: RefreshPivotCacheOptions,
): RefreshPivotCacheResult {
  if (!wb.capabilities.pivotTableMutate) return { ok: false, reason: 'unsupported' };
  if (!writePivotCacheWorksheetSource(wb, opts.cacheId, opts.source)) {
    return { ok: false, reason: 'engine-failed', step: 'cache-source' };
  }
  const refused = wb.refreshPivotCache(opts.cacheId);
  return refused === null ? { ok: true, cacheId: opts.cacheId } : refreshFailure(refused);
}

export function refreshPivotTableFromRange(
  wb: WorkbookHandle,
  opts: RefreshPivotTableOptions,
): RefreshPivotCacheResult {
  if (!wb.capabilities.pivotTableMutate) return { ok: false, reason: 'unsupported' };
  const cacheId = wb.pivotTableCacheId(opts.sheet, opts.pivotIndex);
  if (cacheId < 0) return { ok: false, reason: 'invalid-pivot' };
  return refreshPivotCacheFromRange(wb, { cacheId, source: opts.source });
}

export function refreshPivotTable(
  wb: WorkbookHandle,
  opts: RefreshPivotTableFromStoredSourceOptions,
): RefreshPivotCacheResult {
  if (!wb.capabilities.pivotTableMutate) return { ok: false, reason: 'unsupported' };
  const cacheId = wb.pivotTableCacheId(opts.sheet, opts.pivotIndex);
  if (cacheId < 0) return { ok: false, reason: 'invalid-pivot' };
  const source = pivotCacheWorksheetSourceRange(wb, cacheId);
  if (!source) return { ok: false, reason: 'invalid-range' };
  return refreshPivotCacheFromRange(wb, { cacheId, source });
}
