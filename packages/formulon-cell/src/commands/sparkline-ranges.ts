import { formatA1Range } from '../engine/address.js';
import { parseRangeRef } from '../engine/range-resolver.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';

/** A range typed into a sparkline dialog, with whether it named its sheet. */
export interface ResolvedSparklineRange extends Range {
  explicitSheet: boolean;
}

const formatSheetRange = (sheetName: string, range: Range): string => {
  const body = formatA1Range(range);
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(sheetName)) return `${sheetName}!${body}`;
  return `'${sheetName.replace(/'/g, "''")}'!${body}`;
};

/** Resolve typed range text against `fallbackSheet`; null when it does not parse or names an
 *  unknown sheet. */
export const resolveSparklineRange = (
  raw: string,
  fallbackSheet: number,
  wb: WorkbookHandle | null,
): ResolvedSparklineRange | null => {
  const parsed = parseRangeRef(raw);
  if (!parsed) return null;
  const sheet =
    parsed.sheetName === null
      ? fallbackSheet
      : wb
        ? (wb.sheetIndexByName(parsed.sheetName) ?? -1)
        : -1;
  if (sheet < 0) return null;
  return {
    sheet,
    r0: parsed.r0,
    c0: parsed.c0,
    r1: parsed.r1,
    c1: parsed.c1,
    explicitSheet: parsed.sheetName !== null,
  };
};

/** The source text of a sparkline hosted on `hostSheet` reading `source`: sheet-qualified when
 *  the data is on another sheet or was typed with one. */
export const sparklineSourceText = (
  source: ResolvedSparklineRange,
  hostSheet: number,
  wb: WorkbookHandle | null,
): string => {
  if (source.explicitSheet || source.sheet !== hostSheet) {
    return formatSheetRange(wb?.sheetName(source.sheet) ?? `Sheet${source.sheet + 1}`, source);
  }
  return formatA1Range(source);
};
