import {
  emptyLoadDiagnostics,
  type FileLossId,
  fileLossEntries,
} from '../engine/file-diagnostics.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';

export type WorkbookSaveFormat = 'xlsx' | 'xlsb';

/** File extension and MIME type a host uses to deliver each container. */
export const WORKBOOK_SAVE_FORMATS: Readonly<
  Record<WorkbookSaveFormat, { readonly extension: string; readonly mimeType: string }>
> = {
  xlsx: {
    extension: '.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  },
  xlsb: {
    extension: '.xlsb',
    mimeType: 'application/vnd.ms-excel.sheet.binary.macroEnabled.12',
  },
};

/** Workbook content the engine's `.xlsb` writer can leave out. */
export type XlsbContentLossId =
  | 'notes'
  | 'threadedComments'
  | 'autoFilters'
  | 'drawings'
  | 'pivotTables'
  | 'slicers'
  | 'tables'
  | 'sparklines'
  | 'protectedRanges'
  | 'otherSettings';

/** One kind of content a save leaves out: an itemized `.xlsb` gap or an
 *  engine save counter. `certain` is false where the engine keeps the content
 *  in some cases (`.xlsb`-sourced parts, encodable references) and reports
 *  only an aggregate count. Counts are never summed across rows. */
export interface WorkbookSaveLoss {
  id: XlsbContentLossId | FileLossId;
  count?: number;
  certain: boolean;
}

export interface WorkbookSaveOutcome {
  format: WorkbookSaveFormat;
  bytes: Uint8Array;
  losses: WorkbookSaveLoss[];
}

const sheetIndexes = (wb: WorkbookHandle): number[] =>
  Array.from({ length: wb.sheetCount }, (_, i) => i);

const sumSheets = (wb: WorkbookHandle, count: (sheet: number) => number): number =>
  sheetIndexes(wb).reduce((total, sheet) => total + count(sheet), 0);

/** Itemize what the `.xlsb` writer deferred. Notes, threaded comments and
 *  filters are always dropped; the other categories only sometimes, so they
 *  are listed as possible losses. `otherSettings` covers deferred units no
 *  present category can account for (page setup, views, workbook protection,
 *  custom styles). Each `bound` is the most units a category can add to the
 *  engine's aggregate `deferredFeatureCount`. */
function xlsbContentLosses(wb: WorkbookHandle, deferred: number): WorkbookSaveLoss[] {
  const certain = [
    { id: 'notes', count: sumSheets(wb, (s) => wb.getComments(s).length) },
    { id: 'threadedComments', count: sumSheets(wb, (s) => wb.getThreadedComments(s).length) },
    { id: 'autoFilters', count: sumSheets(wb, (s) => (wb.getAutoFilter(s) ? 1 : 0)) },
  ] as const;
  const possible = [
    {
      id: 'drawings',
      count: sumSheets(wb, (s) => wb.listDrawingObjects(s).length),
      bound: sumSheets(wb, (s) => (wb.listDrawingObjects(s).length > 0 ? 1 : 0)),
    },
    { id: 'pivotTables', count: wb.getPivotTables().length },
    { id: 'slicers', count: sumSheets(wb, (s) => wb.listSlicers(s).length) },
    { id: 'tables', count: wb.getTables().length },
    { id: 'sparklines', count: sumSheets(wb, (s) => wb.listSparklineGroups(s).length) },
    { id: 'protectedRanges', count: sumSheets(wb, (s) => wb.getProtectedRanges(s).length) },
  ] as const;
  const rows: WorkbookSaveLoss[] = [];
  let explained = 0;
  for (const row of certain) {
    if (row.count === 0) continue;
    rows.push({ id: row.id, count: row.count, certain: true });
    explained += row.count;
  }
  for (const row of possible) {
    if (row.count === 0) continue;
    rows.push({ id: row.id, count: row.count, certain: false });
    explained += 'bound' in row ? row.bound : row.count;
  }
  if (deferred > explained) rows.push({ id: 'otherSettings', certain: false });
  return rows;
}

/** Write `wb` as `format` without delivering it, reporting what the file
 *  leaves out so a host can warn before handing the bytes to the user. */
export function writeWorkbook(wb: WorkbookHandle, format: WorkbookSaveFormat): WorkbookSaveOutcome {
  const { bytes, diagnostics } = wb.saveWithDiagnostics(format);
  const engineLosses = fileLossEntries(emptyLoadDiagnostics(), diagnostics).map(
    ({ id, count }): WorkbookSaveLoss => ({ id, count, certain: true }),
  );
  const deferred = diagnostics.deferredFeatureCount;
  const contentLosses = format === 'xlsb' && deferred > 0 ? xlsbContentLosses(wb, deferred) : [];
  return { format, bytes, losses: [...contentLosses, ...engineLosses] };
}
