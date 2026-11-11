import { formatA1Range, quoteSheetName } from '../engine/address.js';
import { parseRangeRef } from '../engine/range-resolver.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';

/** A range typed into a sparkline dialog, with whether it named its sheet. */
export interface ResolvedSparklineRange extends Range {
  explicitSheet: boolean;
}

/** Most cells a typed sparkline range may cover: one full-width row. An unmeasured bound that
 *  keeps a pasted whole-sheet range from being read cell by cell on every paint. */
export const MAX_SPARKLINE_RANGE_CELLS = 16_384;

/** Resolve typed range text against `fallbackSheet`; null when it does not parse, names an
 *  unknown sheet or covers more than `MAX_SPARKLINE_RANGE_CELLS` cells. */
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
  if ((parsed.r1 - parsed.r0 + 1) * (parsed.c1 - parsed.c0 + 1) > MAX_SPARKLINE_RANGE_CELLS) {
    return null;
  }
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
    const name = wb?.sheetName(source.sheet) ?? `Sheet${source.sheet + 1}`;
    return `${quoteSheetName(name)}!${formatA1Range(source)}`;
  }
  return formatA1Range(source);
};
