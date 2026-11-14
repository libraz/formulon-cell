import { DrawingObjectKind, FormulonError, StatusCode } from '@libraz/formulon';
import {
  emptyLoadDiagnostics,
  type FileLossId,
  fileLossEntries,
} from '../engine/file-diagnostics.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { EngineDrawingObject } from '../engine/workbook-handle-annotations.js';

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

/** Workbook content the engine's `.xlsb` writer can leave out.
 *  `notes` and `threadedComments` are no longer reported: the writer now writes them.
 *  `sparklines` and `protectedRanges` the writer refuses instead, so a save holding one
 *  fails with the `.xlsx` offer (see {@link xlsbUnwritableOf}) rather than listing it. */
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

const isDeferredDrawing = (object: EngineDrawingObject): boolean =>
  object.kind === DrawingObjectKind.GraphicFrame || object.kind === DrawingObjectKind.Other;

const sheetIndexes = (wb: WorkbookHandle): number[] =>
  Array.from({ length: wb.sheetCount }, (_, i) => i);

const sumSheets = (wb: WorkbookHandle, count: (sheet: number) => number): number =>
  sheetIndexes(wb).reduce((total, sheet) => total + count(sheet), 0);

/** Itemize what the `.xlsb` writer deferred. Notes and threaded comments are
 *  always written. The rest are written except in some cases (a filter
 *  criterion without the bounds the desktop computed, a drawing holding
 *  another graphic frame, a PivotTable or table the format cannot hold, a
 *  slicer read from `.xlsx`), so each is listed as a possible loss.
 *  `otherSettings` covers deferred units no present category can account for
 *  (print settings, views, default sizes, workbook protection, styles). Each
 *  `bound` is the most units a category can add to the engine's aggregate
 *  `deferredFeatureCount`. */
function xlsbContentLosses(wb: WorkbookHandle, deferred: number): WorkbookSaveLoss[] {
  const possible = [
    { id: 'autoFilters', count: sumSheets(wb, (s) => (wb.getAutoFilter(s) ? 1 : 0)) },
    {
      id: 'drawings',
      count: sumSheets(wb, (s) => wb.listDrawingObjects(s).length),
      // Pictures, charts and shapes are written; only a drawing holding another
      // graphic frame (SmartArt, ...) is deferred.
      bound: sumSheets(wb, (s) => (wb.listDrawingObjects(s).some(isDeferredDrawing) ? 1 : 0)),
    },
    { id: 'pivotTables', count: wb.getPivotTables().length },
    { id: 'slicers', count: sumSheets(wb, (s) => wb.listSlicers(s).length) },
    { id: 'tables', count: wb.getTables().length },
    // Kept for the published loss ids: a book holding either refuses the .xlsb save first.
    { id: 'sparklines', count: sumSheets(wb, (s) => wb.listSparklineGroups(s).length) },
    { id: 'protectedRanges', count: sumSheets(wb, (s) => wb.getProtectedRanges(s).length) },
  ] as const;
  const rows: WorkbookSaveLoss[] = [];
  let explained = 0;
  for (const row of possible) {
    if (row.count === 0) continue;
    rows.push({ id: row.id, count: row.count, certain: false });
    explained += 'bound' in row ? row.bound : row.count;
  }
  if (deferred > explained) rows.push({ id: 'otherSettings', certain: false });
  return rows;
}

/** A feature the `.xlsb` writer refuses outright: the save writes nothing. */
export type XlsbUnwritableId = 'protectedRanges' | 'sparklines' | 'dataBars' | 'unknownFeature';

/** The refused feature and, for a sheet-scoped one, its sheet index. */
export interface XlsbUnwritable {
  id: XlsbUnwritableId;
  sheet: number | null;
}

/** The engine's `kind=` context values, by the feature they name. */
const UNWRITABLE_KINDS: Readonly<Record<string, XlsbUnwritableId>> = {
  protected_range_descriptor: 'protectedRanges',
  protected_range_rects: 'protectedRanges',
  protected_range_sqref: 'protectedRanges',
  protected_range_title: 'protectedRanges',
  sparklines: 'sparklines',
  data_bar: 'dataBars',
};

/** The feature an `.xlsb` save failed on, or null when `error` is not that refusal. */
export function xlsbUnwritableOf(error: unknown): XlsbUnwritable | null {
  if (!(error instanceof FormulonError) || error.code !== StatusCode.IoFeatureNotWritable) {
    return null;
  }
  const kind = /(?:^|\s)kind=(\S+)/.exec(error.context)?.[1] ?? '';
  const sheet = /(?:^|\s)sheet_index=(\d+)/.exec(error.context)?.[1];
  return {
    id: Object.hasOwn(UNWRITABLE_KINDS, kind)
      ? (UNWRITABLE_KINDS[kind] as XlsbUnwritableId)
      : 'unknownFeature',
    sheet: sheet === undefined ? null : Number(sheet),
  };
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
