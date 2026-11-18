import { CellShift, type ConditionalFormat, type ConditionalFormatPatch } from '@libraz/formulon';
import { addrKey, MAX_COL, MAX_ROW, parseAddrKey } from '../engine/address.js';
import { hydrateThreadedCommentsFromEngine } from '../engine/annotations-sync.js';
import {
  type ActiveAxisFormatSnapshot,
  captureActiveAxisFormatOverrideForStructure,
  restoreActiveAxisFormatOverrideForStructure,
} from '../engine/axis-format-sync.js';
import { hydrateConditionalRulesFromEngine } from '../engine/cf-sync.js';
import { hydrateCommentsAndHyperlinksFromEngine } from '../engine/format-sync.js';
import { hydrateMergesFromEngine } from '../engine/merges-sync.js';
import {
  captureSyncedFormatKeysForStructure,
  restoreSyncedFormatKeysForStructure,
} from '../engine/synced-format-keys.js';
import { hydrateTableOverlaysFromEngine } from '../engine/table-sync.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import { hydrateValidationsFromEngine } from '../engine/validation-sync.js';
import {
  type ActiveWholeSheetFontSnapshot,
  captureActiveWholeSheetFontOverrideForStructure,
  restoreActiveWholeSheetFontOverrideForStructure,
} from '../engine/whole-sheet-font-sync.js';
import type { CellPatch, WorkbookHandle } from '../engine/workbook-handle.js';
import { rangeContainsAddr, rangesIntersect } from '../store/selection-geometry.js';
import {
  type CellFormat,
  type ConditionalRule,
  mutators,
  type RangeValidation,
  type SpreadsheetStore,
  type State,
} from '../store/store.js';
import { listComments } from './comment.js';
import { blockDataTableShift } from './data-table-guard.js';
import { tableForCell } from './format-as-table.js';
import type { History } from './history.js';
import { blockPivotEdit } from './pivot-guard.js';
import { blockedByProtection } from './protection.js';
import { deleteCols, deleteVisibleRows, insertCols, insertRows } from './structure.js';

export type InsertCellsDirection = 'down' | 'right';
export type DeleteCellsDirection = 'up' | 'left';

export interface CellRecord {
  addr: Addr;
  value: CellValue;
  formula: string | null;
}

interface EngineCommentRecord {
  addr: Addr;
  author: string | null;
  text: string;
  visible: boolean;
}

export interface ShiftReplayCheckpoint {
  readonly store: State;
  readonly activeAxis: ActiveAxisFormatSnapshot | null;
  readonly activeWholeSheetFont: ActiveWholeSheetFontSnapshot | null;
  readonly syncedFormatKeys: ReadonlySet<string> | null;
  readonly cells: readonly CellRecord[];
  readonly externalFormulas: readonly CellRecord[];
  readonly style: Uint8Array | null;
  readonly comments: readonly EngineCommentRecord[];
  readonly merges: readonly Range[];
}

const cloneStoreState = (store: SpreadsheetStore): State =>
  structuredClone(store.getState()) as State;

export const captureShiftReplayCheckpoint = (
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  affected: Range,
): ShiftReplayCheckpoint => {
  const style =
    wb.capabilities.rangeStyleSnapshot === true ? wb.snapshotRangeStyle(affected) : null;
  if (wb.capabilities.rangeStyleSnapshot === true && style === null) {
    throw new Error('cell shift: native range style replay checkpoint failed');
  }
  return {
    store: cloneStoreState(store),
    activeAxis: captureActiveAxisFormatOverrideForStructure(wb, affected.sheet),
    activeWholeSheetFont: captureActiveWholeSheetFontOverrideForStructure(wb, affected.sheet),
    syncedFormatKeys: captureSyncedFormatKeysForStructure(wb, affected.sheet),
    cells: collectSheetCells(wb, affected.sheet),
    externalFormulas: collectExternalFormulaCells(wb, affected.sheet),
    style,
    comments: collectEngineComments(
      wb,
      affected,
      listComments(store.getState(), affected.sheet).map((comment) => comment.addr),
    ),
    merges: wb.capabilities.merges
      ? wb.getMerges(affected.sheet).map((merge) => ({ ...merge }))
      : [],
  };
};

export const restoreShiftReplayCheckpoint = (
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  affected: Range,
  checkpoint: ShiftReplayCheckpoint,
): void => {
  wb.withEngineSyncMuted(() => {
    store.setState(() => structuredClone(checkpoint.store) as State);
    restoreActiveAxisFormatOverrideForStructure(wb, affected.sheet, checkpoint.activeAxis);
    restoreActiveWholeSheetFontOverrideForStructure(
      wb,
      affected.sheet,
      checkpoint.activeWholeSheetFont,
    );
    restoreSyncedFormatKeysForStructure(wb, affected.sheet, checkpoint.syncedFormatKeys);
    restoreCheckpointCells(wb, affected.sheet, checkpoint.cells, checkpoint.externalFormulas);
    if (checkpoint.style !== null && !wb.restoreRangeStyle(affected.sheet, checkpoint.style)) {
      throw new Error('cell shift: failed to restore transaction style');
    }
    if (wb.capabilities.comments) {
      const currentComments = collectEngineComments(
        wb,
        affected,
        listComments(store.getState(), affected.sheet).map((comment) => comment.addr),
      );
      for (const comment of currentComments) {
        wb.setCommentEntry(comment.addr.sheet, comment.addr.row, comment.addr.col, '', '');
      }
      for (const comment of checkpoint.comments) {
        wb.setCommentEntry(
          comment.addr.sheet,
          comment.addr.row,
          comment.addr.col,
          comment.author,
          comment.text,
          comment.visible,
        );
      }
    }
    if (wb.capabilities.merges) {
      if (!wb.engineClearMerges(affected.sheet)) {
        throw new Error('cell shift: failed to restore merges');
      }
      for (const merge of checkpoint.merges) {
        if (!wb.engineAddMerge(affected.sheet, merge)) {
          throw new Error('cell shift: failed to restore merge');
        }
      }
    }
  });
};

export function insertCells(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  range: Range,
  direction: InsertCellsDirection,
): boolean {
  const lines = wholeLineEdit(store, range);
  if (lines) {
    return lines.axis === 'row'
      ? insertRows(store, wb, history, lines.at, lines.count)
      : insertCols(store, wb, history, lines.at, lines.count);
  }
  if (blockedByProtection(store, range.sheet, 'insert cells')) return false;
  return shiftCellsViaEngine(store, wb, history, range, direction);
}

/** Table Rows Above: insert as many rows as `range` spans, within the table's columns only. */
export function insertTableRows(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  range: Range,
): boolean {
  const table = tableForCell(store.getState().tables.tables, range.sheet, range.r0, range.c0);
  if (!table || (table.showHeader && range.r0 === table.range.r0)) return false;
  const index = table.name ? wb.tableIndexByName(table.name) : null;
  if (index === null || blockedByProtection(store, range.sheet, 'insert cells')) return false;
  const { c0, c1 } = table.range;
  const first = range.r0 - table.range.r0 - (table.showHeader ? 1 : 0);
  const last = first + range.r1 - range.r0;
  return shiftCellsViaEngine(store, wb, history, { ...range, c0, c1 }, 'down', () =>
    wb.insertTableRows(index, first, last),
  );
}

/** Table Columns to the Left: shift the table's rows right by as many columns as `range` spans. */
export function insertTableColumns(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  range: Range,
): boolean {
  const table = tableForCell(store.getState().tables.tables, range.sheet, range.r0, range.c0);
  if (!table) return false;
  const index = table.name ? wb.tableIndexByName(table.name) : null;
  if (index === null || blockedByProtection(store, range.sheet, 'insert cells')) return false;
  const { r0, r1 } = table.range;
  const first = range.c0 - table.range.c0;
  const last = first + range.c1 - range.c0;
  return shiftCellsViaEngine(store, wb, history, { ...range, r0, r1 }, 'right', () =>
    wb.insertTableColumns(index, first, last),
  );
}

export function deleteCells(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  range: Range,
  direction: DeleteCellsDirection,
): boolean {
  const lines = wholeLineEdit(store, range);
  if (lines) {
    return lines.axis === 'row'
      ? deleteVisibleRows(store, wb, history, lines.at, lines.count)
      : deleteCols(store, wb, history, lines.at, lines.count);
  }
  if (blockedByProtection(store, range.sheet, 'delete cells')) return false;
  return shiftCellsViaEngine(store, wb, history, range, direction);
}

/** A range of whole rows or whole columns on the active sheet. Excel ignores
 *  the shift direction for it and edits the rows or columns themselves. */
export function wholeLineEdit(
  store: SpreadsheetStore,
  range: Range,
): { axis: 'row' | 'col'; at: number; count: number } | null {
  if (range.sheet !== store.getState().data.sheetIndex) return null;
  const fullRows = range.c0 === 0 && range.c1 === MAX_COL;
  const fullCols = range.r0 === 0 && range.r1 === MAX_ROW;
  if (fullRows === fullCols) return null;
  return fullRows
    ? { axis: 'row', at: range.r0, count: range.r1 - range.r0 + 1 }
    : { axis: 'col', at: range.c0, count: range.c1 - range.c0 + 1 };
}

const ENGINE_SHIFT: Readonly<Record<InsertCellsDirection | DeleteCellsDirection, CellShift>> = {
  down: CellShift.Down,
  right: CellShift.Right,
  up: CellShift.Up,
  left: CellShift.Left,
};

const INVERSE_SHIFT: Readonly<Record<CellShift, CellShift>> = {
  [CellShift.Down]: CellShift.Up,
  [CellShift.Up]: CellShift.Down,
  [CellShift.Right]: CellShift.Left,
  [CellShift.Left]: CellShift.Right,
};

/** The store slices a cell shift rewrites. */
interface ShiftStoreSlices {
  readonly format: State['format'];
  readonly merges: State['merges'];
  readonly conditional: State['conditional'];
  readonly tables: State['tables'];
  readonly threadedComments: State['threadedComments'];
}

const captureShiftStoreSlices = (state: State): ShiftStoreSlices =>
  structuredClone({
    format: state.format,
    merges: state.merges,
    conditional: state.conditional,
    tables: state.tables,
    threadedComments: state.threadedComments,
  });

/** What the engine cannot rebuild by shifting back: the sheet's cells and
 *  formulas everywhere, the band's styles, notes, merges, conditional formats,
 *  validations, defined names and hyperlinks as they stood. */
interface EngineShiftCheckpoint {
  readonly cells: readonly CellRecord[];
  readonly externalFormulas: readonly CellRecord[];
  readonly style: Uint8Array | null;
  readonly comments: readonly EngineCommentRecord[];
  readonly merges: readonly Range[];
  readonly conditionalFormats: ReturnType<WorkbookHandle['getConditionalFormats']>;
  readonly validations: ReturnType<WorkbookHandle['getValidationRules']>;
  readonly definedNames: ReturnType<typeof listNames>;
  readonly hyperlinks: ReturnType<WorkbookHandle['getHyperlinksFull']>;
}

const listNames = (wb: WorkbookHandle) => [...wb.definedNames()];

/** Null when the band's styles cannot be captured; the shift is then refused. */
function captureEngineShiftCheckpoint(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  band: Range,
): EngineShiftCheckpoint | null {
  const sheet = band.sheet;
  const style = wb.capabilities.rangeStyleSnapshot === true ? wb.snapshotRangeStyle(band) : null;
  if (wb.capabilities.rangeStyleSnapshot === true && style === null) return null;
  return {
    cells: collectSheetCells(wb, sheet),
    externalFormulas: collectExternalFormulaCells(wb, sheet),
    style,
    comments: collectEngineComments(
      wb,
      band,
      listComments(store.getState(), sheet).map((comment) => comment.addr),
    ),
    merges: wb.capabilities.merges ? wb.getMerges(sheet).map((merge) => ({ ...merge })) : [],
    conditionalFormats: wb.getConditionalFormats(sheet),
    validations: wb.getValidationRules(sheet),
    definedNames: listNames(wb),
    hyperlinks: wb.getHyperlinksFull(sheet),
  };
}

/** Put the engine back to `checkpoint` after the inverse shift moved the band back. */
function restoreEngineShiftCheckpoint(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  band: Range,
  checkpoint: EngineShiftCheckpoint,
): void {
  const sheet = band.sheet;
  restoreCheckpointCells(wb, sheet, checkpoint.cells, checkpoint.externalFormulas);
  if (checkpoint.style !== null && !wb.restoreRangeStyle(sheet, checkpoint.style)) {
    throw new Error('cell shift: failed to restore band styles');
  }
  if (wb.capabilities.comments) {
    const current = collectEngineComments(
      wb,
      band,
      listComments(store.getState(), sheet).map((comment) => comment.addr),
    );
    for (const comment of current)
      wb.setCommentEntry(sheet, comment.addr.row, comment.addr.col, '', '');
    for (const comment of checkpoint.comments) {
      wb.setCommentEntry(
        sheet,
        comment.addr.row,
        comment.addr.col,
        comment.author,
        comment.text,
        comment.visible,
      );
    }
  }
  if (wb.capabilities.merges) {
    if (!wb.engineClearMerges(sheet)) throw new Error('cell shift: failed to restore merges');
    for (const merge of checkpoint.merges) {
      if (!wb.engineAddMerge(sheet, merge)) throw new Error('cell shift: failed to restore merge');
    }
  }
  if (wb.capabilities.conditionalFormatMutate) {
    restoreRulesById(
      {
        list: () => wb.getConditionalFormats(sheet),
        removeAt: (index) => wb.removeConditionalFormatAt(sheet, index),
        update: (rule, current) =>
          wb.updateConditionalFormat(sheet, rule.id, conditionalFormatPatch(current, rule)),
        add: (rule) => wb.addConditionalFormat(sheet, rule) >= 0,
      },
      checkpoint.conditionalFormats,
      'conditional format',
    );
  }
  if (wb.capabilities.dataValidation) {
    restoreRulesById(
      {
        list: () => wb.getValidationRules(sheet),
        removeAt: (index) => wb.removeValidationAt(sheet, index),
        update: (rule) => wb.updateValidationRule(sheet, rule.id, rule),
        add: (rule) => wb.addValidationRule(sheet, rule),
      },
      checkpoint.validations,
      'validation',
    );
  }
  const names = new Map(listNames(wb).map((name) => [`${name.localSheetId}:${name.name}`, name]));
  for (const name of checkpoint.definedNames) {
    const current = names.get(`${name.localSheetId}:${name.name}`);
    if (current?.formula !== name.formula) {
      wb.setDefinedNameEntry(name.name, name.formula, name.localSheetId);
    }
  }
  if (checkpoint.hyperlinks && wb.capabilities.hyperlinks) {
    wb.clearHyperlinks(sheet);
    for (const link of checkpoint.hyperlinks) {
      wb.addHyperlinkRange(
        sheet,
        link.range.firstRow,
        link.range.firstCol,
        link.range.lastRow,
        link.range.lastCol,
        link.target ?? '',
        link.display ?? '',
        link.tooltip ?? '',
        link.location ?? '',
      );
    }
  }
  wb.recalcAuto();
}

/** Engine writes over one sheet's id-keyed rules (conditional formats or validations). */
interface RuleWriter<R extends { readonly id: string }> {
  list(): readonly R[];
  removeAt(index: number): boolean;
  update(rule: R, current: R): boolean;
  add(rule: R): boolean;
}

/** Bring the sheet's rules to `target` by id: rules missing from it are removed,
 *  changed ones updated in place, gone ones added back. A refused write puts the
 *  rules back as they were, then throws. */
function restoreRulesById<R extends { readonly id: string }>(
  writer: RuleWriter<R>,
  target: readonly R[],
  what: string,
): void {
  const before = writer.list();
  if (reconcileRules(writer, before, target)) return;
  reconcileRules(writer, writer.list(), before);
  throw new Error(`cell shift: failed to restore a ${what}`);
}

/** False at the first write the engine refuses. */
function reconcileRules<R extends { readonly id: string }>(
  writer: RuleWriter<R>,
  current: readonly R[],
  target: readonly R[],
): boolean {
  const wanted = new Set(target.map((rule) => rule.id));
  const present = new Map(current.map((rule) => [rule.id, rule]));
  // Highest index first so the earlier positions stay valid.
  const removals = current.flatMap((rule, index) => (wanted.has(rule.id) ? [] : [index])).reverse();
  const updates = target.flatMap((rule) => {
    const now = present.get(rule.id);
    return now && JSON.stringify(now) !== JSON.stringify(rule) ? [{ rule, now }] : [];
  });
  const additions = target.filter((rule) => !present.has(rule.id));
  return (
    removals.every((index) => writer.removeAt(index)) &&
    updates.every(({ rule, now }) => writer.update(rule, now)) &&
    additions.every((rule) => writer.add(rule))
  );
}

/** The keys of a CF update: a rule's type and id never change. */
const CF_PATCH_KEYS = [
  'ranges',
  'priority',
  'stopIfTrue',
  'dxfIndex',
  'operator',
  'formula1',
  'formula2',
  'colorScale',
  'dataBar',
  'iconSet',
  'rank',
  'percent',
  'bottom',
  'aboveAverage',
  'equalAverage',
  'stdDev',
  'text',
  'timePeriod',
] as const satisfies readonly (keyof ConditionalFormat)[];

/** The keys of `target` that differ from `current`; a key the rule's kind does
 *  not use reads null on both and is left out, as the engine requires. */
function conditionalFormatPatch(
  current: ConditionalFormat,
  target: ConditionalFormat,
): ConditionalFormatPatch {
  const patch: Record<string, unknown> = {};
  for (const key of CF_PATCH_KEYS) {
    if (JSON.stringify(current[key]) !== JSON.stringify(target[key])) patch[key] = target[key];
  }
  return patch as ConditionalFormatPatch;
}

/** Bring the store in line with an engine shift of `range`. The engine is
 *  read back for cells, merges, imported conditional formats, tables and
 *  threaded comments. Per-cell formats are shifted in the store: the engine
 *  cannot yet enumerate format-only cells. Session conditional formats, which
 *  the engine has not been given yet, shift with their cells. */
function mirrorEngineShift(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  band: Range,
  axis: InsertCellsDirection,
  delta: number,
): void {
  const sheet = band.sheet;
  wb.withEngineSyncMuted(() => {
    store.setState((s) => {
      const shifted = shiftFormatMap(s.format.formats, band, axis, delta);
      const formats =
        delta > 0
          ? inheritInsertedValidations(s.format.formats, shifted, band, axis, delta)
          : shifted;
      const rangeValidations = shiftRangeValidationsForBand(
        s.format.rangeValidations,
        band,
        axis,
        delta,
      );
      const rules = s.conditional.rules.flatMap<ConditionalRule>((rule) => {
        if (rule.range.sheet !== sheet || rule.engineId !== undefined) return [rule];
        return rangePiecesAfterBandShift(rule.range, band, axis, delta, 'inside').map((range) => ({
          ...rule,
          range,
        }));
      });
      return {
        ...s,
        format: { ...s.format, formats, rangeValidations },
        conditional: { rules },
      };
    });
    if (sheet === store.getState().data.sheetIndex) mutators.replaceCells(store, wb.cells(sheet));
    hydrateCommentsAndHyperlinksFromEngine(wb, store, sheet);
    hydrateMergesFromEngine(wb, store, sheet);
    hydrateConditionalRulesFromEngine(wb, store, sheet);
    hydrateValidationsFromEngine(wb, store, sheet);
    hydrateTableOverlaysFromEngine(wb, store);
    hydrateThreadedCommentsFromEngine(wb, store, sheet);
  });
}

/** Move the store's per-cell formats and range validations of `band` by
 *  `delta` rows after the engine shifted the band's cells (a totals row
 *  turned on or off). The engine cannot yet enumerate format-only cells. */
export function mirrorBandFormatsDown(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  band: Range,
  delta: number,
): void {
  wb.withEngineSyncMuted(() => {
    store.setState((s) => ({
      ...s,
      format: {
        ...s.format,
        formats: shiftFormatMap(s.format.formats, band, 'down', delta),
        rangeValidations: shiftRangeValidationsForBand(
          s.format.rangeValidations,
          band,
          'down',
          delta,
        ),
      },
    }));
  });
}

function restoreShiftStoreSlices(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  sheet: number,
  slices: ShiftStoreSlices,
): void {
  wb.withEngineSyncMuted(() => {
    store.setState((s) => ({ ...s, ...structuredClone(slices) }));
    if (sheet === store.getState().data.sheetIndex) mutators.replaceCells(store, wb.cells(sheet));
  });
}

/** Insert or delete the cells of `range` through the engine, which moves the
 *  band's cells, adjusts every reference and splits conditional formats,
 *  validations and merges. `apply` performs the forward edit (a table insert
 *  in place of the plain shift). One history step; undo shifts back and
 *  restores what the shift removed. */
export function shiftCellsViaEngine(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  range: Range,
  direction: InsertCellsDirection | DeleteCellsDirection,
  apply: () => boolean = () => wb.engineShiftCells(range, ENGINE_SHIFT[direction]),
): boolean {
  // Shifting cells is an engine operation; the stub engine cannot do it.
  if (!wb.capabilities.insertDeleteRowsCols) return false;
  const sheet = range.sheet;
  const axis: InsertCellsDirection = direction === 'down' || direction === 'up' ? 'down' : 'right';
  const band: Range =
    axis === 'down'
      ? { sheet, r0: range.r0, c0: range.c0, r1: MAX_ROW, c1: range.c1 }
      : { sheet, r0: range.r0, c0: range.c0, r1: range.r1, c1: MAX_COL };
  const span = axis === 'down' ? range.r1 - range.r0 + 1 : range.c1 - range.c0 + 1;
  const delta = direction === 'down' || direction === 'right' ? span : -span;
  const shift = ENGINE_SHIFT[direction];
  const record = history !== null && !history.isReplaying();
  const beforeEngine = record ? captureEngineShiftCheckpoint(wb, store, band) : null;
  if (record && beforeEngine === null) return false;
  const beforeStore = record ? captureShiftStoreSlices(store.getState()) : null;
  if (!apply()) {
    explainRefusedShift(wb, band);
    return false;
  }
  wb.recalcAuto();
  mirrorEngineShift(store, wb, band, axis, delta);
  if (record && beforeEngine && beforeStore) {
    const afterStore = captureShiftStoreSlices(store.getState());
    history.push({
      undo: () => {
        if (!wb.engineShiftCells(range, INVERSE_SHIFT[shift])) {
          throw new Error('cell shift: the engine refused to shift back');
        }
        restoreEngineShiftCheckpoint(wb, store, band, beforeEngine);
        restoreShiftStoreSlices(store, wb, sheet, beforeStore);
      },
      redo: () => {
        if (!apply()) throw new Error('cell shift: redo was refused');
        wb.recalcAuto();
        restoreShiftStoreSlices(store, wb, sheet, afterStore);
      },
    });
  }
  return true;
}

/** Tell the user why the engine refused a shift: a PivotTable or data table in
 *  the way has its own alert; anything else (a table, an array formula, content
 *  pushed off the grid) is logged. */
function explainRefusedShift(wb: WorkbookHandle, band: Range): void {
  if (blockPivotEdit(wb, band)) return;
  if (blockDataTableShift(wb, band)) return;
  // eslint-disable-next-line no-console
  console.warn('formulon-cell: cell shift blocked by a table, array formula or the grid edge');
}

function collectSheetCells(wb: WorkbookHandle, sheet: number): CellRecord[] {
  return Array.from(wb.physicalCells(sheet)).map<CellRecord>((c) => ({
    addr: c.addr,
    value: c.value,
    formula: c.formula,
  }));
}

function collectExternalFormulaCells(wb: WorkbookHandle, editedSheet: number): CellRecord[] {
  const formulas: CellRecord[] = [];
  for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) {
    if (sheet === editedSheet) continue;
    for (const cell of wb.physicalCells(sheet)) {
      if (!cell.formula) continue;
      formulas.push({ addr: cell.addr, value: cell.value, formula: cell.formula });
    }
  }
  return formulas;
}

/** Put the cells of `sheet` and the formulas of the other sheets back as one
 *  engine transaction: a refused write leaves every cell as it was. */
function restoreCheckpointCells(
  wb: WorkbookHandle,
  sheet: number,
  cells: readonly CellRecord[],
  externalFormulas: readonly CellRecord[],
): void {
  const kept = new Set(cells.map((cell) => addrKey(cell.addr)));
  const blanks = Array.from(wb.physicalCells(sheet)).flatMap<CellPatch>((cell) =>
    kept.has(addrKey(cell.addr))
      ? []
      : [{ addr: cell.addr, value: { kind: 'blank' }, formula: null }],
  );
  wb.applyCellPatchAtomic([...cells, ...blanks, ...externalFormulas]);
}

function collectEngineComments(
  wb: WorkbookHandle,
  affected: Range,
  storeCommentAddresses: readonly Addr[],
): EngineCommentRecord[] {
  const entries = wb.capabilities.commentsEnumerable
    ? wb.getComments(affected.sheet).map((comment) => ({
        addr: { sheet: affected.sheet, row: comment.row, col: comment.col },
        author: comment.author,
        text: comment.text,
        visible: comment.visible === true,
      }))
    : storeCommentAddresses.flatMap((addr) => {
        const comment = wb.getComment(addr.sheet, addr.row, addr.col);
        return comment
          ? [
              {
                addr,
                author: comment.author,
                text: comment.text,
                visible: comment.visible === true,
              },
            ]
          : [];
      });
  const unique = new Map<string, EngineCommentRecord>();
  for (const entry of entries) {
    if (rangeContainsAddr(affected, entry.addr)) unique.set(addrKey(entry.addr), entry);
  }
  return [...unique.values()];
}

/** Whether every merge touching the shift band moves whole; false when a
 *  merge straddles the band edge and would be split. */
export function canShiftMerges(state: State, affected: Range, axis: InsertCellsDirection): boolean {
  for (const merge of state.merges.byAnchor.values()) {
    if (merge.sheet !== affected.sheet) continue;
    if (!rangesIntersect(merge, affected)) continue;
    if (!mergeIntersectsShiftBand(merge, affected, axis)) continue;
    if (!mergeFitsBand(merge, affected, axis)) return false;
  }
  return true;
}

function mergeFitsBand(merge: Range, affected: Range, axis: InsertCellsDirection): boolean {
  return axis === 'down'
    ? merge.c0 >= affected.c0 && merge.c1 <= affected.c1 && merge.r0 >= affected.r0
    : merge.r0 >= affected.r0 && merge.r1 <= affected.r1 && merge.c0 >= affected.c0;
}

const transposeRange = (r: Range): Range => ({ ...r, r0: r.c0, c0: r.r0, r1: r.c1, c1: r.r1 });

/** Areas a range covers once the cells of `affected` move by `delta`. Cells
 *  outside the band stay; band cells move, and deleted ones vanish. Inserted
 *  cells join a range that has the cell before the band (`edge`) or both
 *  neighbours of the insertion point (`inside`). */
function rangePiecesAfterBandShift(
  range: Range,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
  inherit: 'inside' | 'edge',
): Range[] {
  if (range.sheet !== affected.sheet) return [range];
  if (axis === 'right') {
    return rangePiecesAfterBandShift(
      transposeRange(range),
      transposeRange(affected),
      'down',
      delta,
      inherit,
    ).map(transposeRange);
  }
  const m0 = Math.max(range.c0, affected.c0);
  const m1 = Math.min(range.c1, affected.c1);
  if (m0 > m1 || range.r1 < affected.r0) return [range];
  const pieces: Range[] = [];
  if (range.c0 < m0) pieces.push({ ...range, c1: m0 - 1 });
  if (range.c1 > m1) pieces.push({ ...range, c0: m1 + 1 });
  if (range.r0 < affected.r0) {
    pieces.push({ ...range, c0: m0, c1: m1, r1: Math.min(range.r1, affected.r0 - 1) });
  }
  const moved = (from: number, to: number): void => {
    const r0 = from + delta;
    const r1 = Math.min(to + delta, MAX_ROW);
    if (r0 <= r1) pieces.push({ ...range, c0: m0, c1: m1, r0, r1 });
  };
  if (delta > 0) {
    moved(Math.max(range.r0, affected.r0), range.r1);
    const above = affected.r0 - 1;
    if (range.r0 <= above && above <= range.r1 && (inherit === 'edge' || range.r1 >= affected.r0)) {
      pieces.push({
        ...range,
        c0: m0,
        c1: m1,
        r0: affected.r0,
        r1: Math.min(affected.r0 + delta - 1, MAX_ROW),
      });
    }
  } else {
    const from = Math.max(range.r0, affected.r0 - delta);
    if (from <= range.r1) moved(from, range.r1);
    else if (inherit === 'inside') {
      // A deleted block that reaches the range's trailing edge leaves the rule whole.
      pieces.push({ ...range, c0: m0, c1: m1, r0: Math.max(range.r0, affected.r0) });
    }
  }
  return coalesceRanges(pieces);
}

/** Merge pieces that share a full edge into single rectangles. */
function coalesceRanges(pieces: Range[]): Range[] {
  const out = pieces.map((p) => ({ ...p }));
  let merged = true;
  while (merged) {
    merged = false;
    for (let i = 0; i < out.length && !merged; i += 1) {
      for (let j = i + 1; j < out.length && !merged; j += 1) {
        const a = out[i] as Range;
        const b = out[j] as Range;
        const sameCols = a.c0 === b.c0 && a.c1 === b.c1;
        const sameRows = a.r0 === b.r0 && a.r1 === b.r1;
        if (sameCols && a.r0 <= b.r1 + 1 && b.r0 <= a.r1 + 1) {
          out[i] = { ...a, r0: Math.min(a.r0, b.r0), r1: Math.max(a.r1, b.r1) };
        } else if (sameRows && a.c0 <= b.c1 + 1 && b.c0 <= a.c1 + 1) {
          out[i] = { ...a, c0: Math.min(a.c0, b.c0), c1: Math.max(a.c1, b.c1) };
        } else continue;
        out.splice(j, 1);
        merged = true;
      }
    }
  }
  return out.sort((a, b) => a.r0 - b.r0 || a.c0 - b.c0);
}

/** Range-stored validations follow the cells of the band; inserted cells take
 *  the rule of the cell above (or to the left). */
function shiftRangeValidationsForBand(
  rules: RangeValidation[] | undefined,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): RangeValidation[] | undefined {
  if (!rules?.some((rule) => rule.ranges.some((range) => range.sheet === affected.sheet))) {
    return rules;
  }
  return rules.flatMap((rule) => {
    const ranges = rule.ranges.flatMap((range) =>
      rangePiecesAfterBandShift(range, affected, axis, delta, 'edge'),
    );
    return ranges.length > 0 ? [{ ...rule, ranges }] : [];
  });
}

/** Inserted cells take the validation of the cell above (or to the left). */
function inheritInsertedValidations(
  before: Map<string, CellFormat>,
  shifted: Map<string, CellFormat>,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): Map<string, CellFormat> {
  let out = shifted;
  const down = axis === 'down';
  const alongStart = down ? affected.r0 : affected.c0;
  const acrossLo = down ? affected.c0 : affected.r0;
  const acrossHi = down ? affected.c1 : affected.r1;
  for (const [key, fmt] of before) {
    if (!fmt.validation) continue;
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== affected.sheet) continue;
    const along = down ? addr.row : addr.col;
    const across = down ? addr.col : addr.row;
    if (along !== alongStart - 1 || across < acrossLo || across > acrossHi) continue;
    if (out === shifted) out = new Map(shifted);
    for (
      let at = alongStart;
      at < alongStart + delta && at <= (down ? MAX_ROW : MAX_COL);
      at += 1
    ) {
      const target = addrKey(down ? { ...addr, row: at } : { ...addr, col: at });
      out.set(target, { ...out.get(target), validation: fmt.validation });
    }
  }
  return out;
}

function shiftFormatMap(
  formats: Map<string, CellFormat>,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): Map<string, CellFormat> {
  const next = new Map<string, CellFormat>();
  for (const [key, fmt] of formats) {
    const addr = parseAddrKey(key);
    if (!addr) {
      next.set(key, fmt);
      continue;
    }
    if (!rangeContainsAddr(affected, addr)) {
      next.set(key, fmt);
      continue;
    }
    const shifted =
      axis === 'down' ? { ...addr, row: addr.row + delta } : { ...addr, col: addr.col + delta };
    if (inShiftTarget(shifted, affected, axis)) next.set(addrKey(shifted), fmt);
  }
  return next;
}

function inShiftTarget(addr: Addr, affected: Range, axis: InsertCellsDirection): boolean {
  if (addr.sheet !== affected.sheet || addr.row < 0 || addr.col < 0) return false;
  if (addr.row > MAX_ROW || addr.col > MAX_COL) return false;
  return axis === 'down'
    ? addr.row >= affected.r0 && addr.col >= affected.c0 && addr.col <= affected.c1
    : addr.col >= affected.c0 && addr.row >= affected.r0 && addr.row <= affected.r1;
}

function mergeIntersectsShiftBand(
  merge: Range,
  affected: Range,
  axis: InsertCellsDirection,
): boolean {
  return axis === 'down'
    ? merge.r1 >= affected.r0 && merge.c1 >= affected.c0 && merge.c0 <= affected.c1
    : merge.c1 >= affected.c0 && merge.r1 >= affected.r0 && merge.r0 <= affected.r1;
}
