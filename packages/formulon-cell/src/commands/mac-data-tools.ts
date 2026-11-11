import { colLetter, formatA1Cell, MAX_COL, MAX_ROW } from '../engine/address.js';
import { parseRangeRef } from '../engine/range-resolver.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import { type CellSnapshot, WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import type { SpreadsheetInstance } from '../mount/types.js';
import { sameAddr } from '../store/direct-format.js';
import { outlineSummaryFor } from '../store/outline-summary.js';
import { rangeArea, rangesIntersect, sameRange } from '../store/selection-geometry.js';
import { getPageSetup, mutators, type SpreadsheetStore } from '../store/store.js';
import {
  type AuthorizedIntentFactory,
  authorizedWrite,
  authorizeWrite,
} from './authorized-write.js';
import { blockDataTableShift } from './data-table-guard.js';
import { rowShiftBand } from './interaction-effects.js';
import { groupRows, ungroupRows } from './outline.js';
import { isSheetProtected } from './protection.js';
import { sameSavedWorkbook } from './saved-workbook-proof.js';
import { recordFormatChange, recordPageSetupChange } from './slice-history.js';
import { deleteRows, insertRows } from './structure.js';

/** Data commands exposed by the Mac ribbon. The limits are
 * deliberately finite because these commands materialize every output cell
 * before committing it. */
export const MAX_CONSOLIDATE_CELLS = 100_000;
export const MAX_SUBTOTAL_ROWS = 100_000;
export const MAX_SUBTOTAL_OUTPUT_CELLS = 100_000;
export const MAX_SUBTOTAL_GROUPS = 1_000;
export const MAX_GOAL_SEEK_ITERATIONS = 100;
/** Desktop default Maximum Change: Goal Seek stops once |formula - target| is below it. */
const GOAL_SEEK_MAX_CHANGE = 0.001;
/** First trial offset, as a fraction of |x0| (1 when x0 is 0). When neither
 * side helps, the search walks down from x0 by this step doubled, up to 2^25. */
const GOAL_SEEK_PROBE = 0.01;
const GOAL_SEEK_PROBE_DOUBLINGS = 25;
/** Secant steps are capped at this fraction of the scale, doubling per step. */
const GOAL_SEEK_STEP_LIMIT = 0.1;
/** Measured on desktop: a search that never improves ends at the step limit
 * 0.1 * 2^48, i.e. after 49 steps past the first trial. */
const GOAL_SEEK_SEARCH_STEPS = 49;

export type MacDataFunction = 'sum' | 'average' | 'count' | 'min' | 'max';

const MAC_DATA_FUNCTIONS: ReadonlySet<MacDataFunction> = new Set([
  'sum',
  'average',
  'count',
  'min',
  'max',
]);

/** Keys of the localized `macData.errors` table. */
export type MacDataErrorCode = keyof Strings['mac']['data']['errors'];

export interface MacDataError {
  readonly status: 'rejected' | 'invalid' | 'unsupported' | 'stale' | 'nonconverged';
  readonly code: MacDataErrorCode;
  /** Detail supplied by the host policy, command pipeline, or a thrown error.
   *  Absent when the code alone describes the failure. */
  readonly reason?: string;
}

/** Localized text for a failure: the host-supplied detail when present,
 *  otherwise the table entry for its code. */
export const describeMacDataError = (strings: Strings, error: MacDataError): string =>
  error.reason ?? strings.mac.data.errors[error.code];

export type MacDataResult<T> =
  | { readonly ok: true; readonly value: T }
  | { readonly ok: false; readonly error: MacDataError };

export interface GoalSeekRequest {
  readonly formulaCell: Addr;
  readonly targetValue: number;
  readonly changingCell: Addr;
  readonly tolerance?: number;
  readonly maxIterations?: number;
}

export interface GoalSeekSolution {
  readonly changingValue: number;
  readonly formulaValue: number;
  readonly iterations: number;
  /** False when the iteration stopped without reaching the target; the last
   *  iterate is still the value the desktop app leaves in the cell. */
  readonly converged: boolean;
}

interface GoalSeekProof {
  readonly sourceBytes: Uint8Array;
  readonly formulaCell: Addr;
  readonly changingCell: Addr;
  readonly targetValue: number;
  readonly tolerance: number | undefined;
  readonly maxIterations: number | undefined;
  readonly solution: GoalSeekSolution;
}

const goalSeekProofs = new WeakMap<GoalSeekSolution, GoalSeekProof>();

export interface ConsolidateRequest {
  readonly sources: readonly string[];
  readonly destination: string;
  readonly function: MacDataFunction;
  /** Match columns across sources by the labels in each source's top row. */
  readonly topRow?: boolean;
  /** Match rows across sources by the labels in each source's left column. */
  readonly leftColumn?: boolean;
}

export interface ConsolidateSolution {
  readonly destination: Range;
  readonly values: readonly CellValue[];
}

export interface SubtotalRequest {
  readonly range: string;
  readonly groupByColumn: number;
  readonly subtotalColumns: readonly number[];
  readonly function: Exclude<MacDataFunction, 'count'> | 'count';
  /** Start every group after the first on a new printed page. */
  readonly pageBreaks?: boolean;
  /** Put each subtotal row below its group and the grand total last; when
   *  false they sit above their data, the grand total under the header. */
  readonly summaryBelow?: boolean;
}

export interface SubtotalSolution {
  readonly range: Range;
  readonly insertedRows: number;
  readonly groups: readonly {
    readonly start: number;
    readonly end: number;
    readonly insertAt: number;
    readonly label: string;
  }[];
}

export interface ParsedMacRange {
  readonly range: Range;
  readonly explicitSheet: boolean;
}

const failure = (
  status: MacDataError['status'],
  code: MacDataErrorCode,
  reason?: string,
): MacDataResult<never> => ({
  ok: false,
  error: reason === undefined ? { status, code } : { status, code, reason },
});

/** The message of a thrown value, when it carries one. */
const thrownReason = (error: unknown): string | undefined =>
  error instanceof Error && error.message ? error.message : undefined;

const success = <T>(value: T): MacDataResult<T> => ({ ok: true, value });

const isFiniteNumber = (value: CellValue): value is { kind: 'number'; value: number } =>
  value.kind === 'number' && Number.isFinite(value.value);

const sameCellValue = (a: CellValue, b: CellValue): boolean => {
  if (a.kind !== b.kind) return false;
  switch (a.kind) {
    case 'blank':
      return true;
    case 'number':
      return b.kind === 'number' && Object.is(a.value, b.value);
    case 'bool':
      return b.kind === 'bool' && a.value === b.value;
    case 'text':
      return b.kind === 'text' && a.value === b.value;
    case 'error':
      return b.kind === 'error' && a.code === b.code && a.text === b.text;
  }
};

const sameSnapshot = (a: CellSnapshot, b: CellSnapshot): boolean =>
  sameAddr(a.addr, b.addr) && a.formula === b.formula && sameCellValue(a.value, b.value);

/** Resolve a user range against the active sheet. `parseRangeRef` intentionally
 * remains the one grammar for A1 references, including quoted sheet names. */
export function parseMacRange(
  wb: WorkbookHandle,
  raw: string,
  fallbackSheet: number,
): ParsedMacRange | null {
  const parsed = parseRangeRef(raw);
  if (!parsed) return null;
  let sheet = fallbackSheet;
  if (parsed.sheetName !== null) {
    sheet = wb.sheetIndexByName(parsed.sheetName) ?? -1;
    if (sheet < 0) return null;
  }
  return {
    explicitSheet: parsed.sheetName !== null,
    range: {
      sheet,
      r0: parsed.r0,
      c0: parsed.c0,
      r1: parsed.r1,
      c1: parsed.c1,
    },
  };
}

const addrRange = (addr: Addr): Range => ({
  sheet: addr.sheet,
  r0: addr.row,
  c0: addr.col,
  r1: addr.row,
  c1: addr.col,
});

const cellAddresses = (range: Range): Addr[] => {
  const out: Addr[] = [];
  for (let row = range.r0; row <= range.r1; row += 1) {
    for (let col = range.c0; col <= range.c1; col += 1) {
      out.push({ sheet: range.sheet, row, col });
    }
  }
  return out;
};

const cellSnapshot = (wb: WorkbookHandle, addr: Addr): CellSnapshot => ({
  addr,
  value: wb.getValue(addr),
  formula: wb.cellFormula(addr),
});

const cellPatchFromSnapshot = (snapshot: CellSnapshot) => ({
  addr: snapshot.addr,
  value: snapshot.value,
  formula: snapshot.formula,
});

const formulaNumber = (wb: WorkbookHandle, addr: Addr): number | null => {
  const value = wb.getValue(addr);
  return isFiniteNumber(value) ? value.value : null;
};

const dataIntent = (
  instance: SpreadsheetInstance,
  operation: 'valueEdit' | 'formulaEdit',
  cells: readonly Addr[],
  commandId: string,
) =>
  instance.commands.canExecute({
    operation,
    origin: 'ribbon',
    commandId,
    effects: [{ kind: 'cells', cells, includesFormula: operation === 'formulaEdit' }],
  });

/** The sheet-level writes of Subtotal: the outline it groups (or ungroups) and, when asked, the
 *  page breaks it adds. */
const subtotalSheetIntents =
  (sheet: number, pageBreaks: boolean): AuthorizedIntentFactory =>
  (origin) => [
    {
      operation: 'resizeRows',
      origin,
      commandId: 'mac.data.subtotal',
      effects: [{ kind: 'sheet', sheet }],
    },
    ...(pageBreaks
      ? [
          {
            operation: 'pageSetup' as const,
            origin,
            commandId: 'mac.data.subtotal',
            effects: [{ kind: 'sheet' as const, sheet }],
          },
        ]
      : []),
  ];

/** Remove All: the subtotal rows deleted below the header (re-inserted on undo) and the outline
 *  cleared. */
const subtotalRemovalIntents =
  (range: Range): AuthorizedIntentFactory =>
  (origin) => [
    {
      operation: origin === 'undo' ? 'insertRows' : 'deleteRows',
      origin,
      commandId: 'mac.data.subtotal',
      effects: [{ kind: 'range', range: rowShiftBand(range.sheet, range.r0 + 1) }],
    },
    ...subtotalSheetIntents(range.sheet, false)(origin),
  ];

/** Replacing subtotals removes the old rows and inserts new ones below the header, writing values,
 *  formulas and formats there; the new rows are known only once the old ones are gone, so the
 *  cells are decided as the band they can land in. */
const subtotalReplaceIntents =
  (range: Range, pageBreaks: boolean): AuthorizedIntentFactory =>
  (origin) => {
    const band = rowShiftBand(range.sheet, range.r0 + 1);
    const cells = { ...band, c0: range.c0, c1: range.c1 };
    return [
      ...(['deleteRows', 'insertRows'] as const).map((operation) => ({
        operation,
        origin,
        commandId: 'mac.data.subtotal',
        effects: [{ kind: 'range' as const, range: band }],
      })),
      ...(['valueEdit', 'formulaEdit'] as const).map((operation) => ({
        operation,
        origin,
        commandId: 'mac.data.subtotal',
        effects: [{ kind: 'range' as const, range: cells }],
      })),
      {
        operation: 'format',
        origin,
        commandId: 'mac.data.subtotal',
        effects: [{ kind: 'workbook' }],
      },
      ...subtotalSheetIntents(range.sheet, pageBreaks)(origin),
    ];
  };

const subtotalIntents = (
  range: Range,
  outputCells: readonly Addr[],
  formulaCells: readonly Addr[],
  pageBreaks: boolean,
): AuthorizedIntentFactory => {
  const cellsEffect = { kind: 'cells' as const, cells: outputCells };
  const formulasEffect = {
    kind: 'cells' as const,
    cells: formulaCells,
    includesFormula: true,
  };
  const formatEffect = { kind: 'workbook' as const };
  // Every subtotal row lands at or below the first data row and shifts the rest of the sheet.
  const shiftedBand = rowShiftBand(range.sheet, range.r0 + 1);
  const structuralEffects = [{ kind: 'range' as const, range: shiftedBand }] as const;
  return (origin) => [
    {
      operation: origin === 'undo' ? 'deleteRows' : 'insertRows',
      origin,
      commandId: 'mac.data.subtotal',
      effects: structuralEffects,
    },
    {
      operation: 'valueEdit',
      origin,
      commandId: 'mac.data.subtotal',
      effects: [cellsEffect],
    },
    {
      operation: 'formulaEdit',
      origin,
      commandId: 'mac.data.subtotal',
      effects: [formulasEffect],
    },
    {
      operation: 'format',
      origin,
      commandId: 'mac.data.subtotal',
      effects: [formatEffect],
    },
    ...subtotalSheetIntents(range.sheet, pageBreaks)(origin),
  ];
};

const parseDestination = (
  wb: WorkbookHandle,
  raw: string,
  fallbackSheet: number,
  rows: number,
  cols: number,
): Range | null => {
  const parsed = parseMacRange(wb, raw, fallbackSheet);
  if (!parsed) return null;
  const area = parsed.range;
  if (rangeArea(area) === 1) {
    const r1 = area.r0 + rows - 1;
    const c1 = area.c0 + cols - 1;
    if (r1 > MAX_ROW || c1 > MAX_COL) return null;
    return { sheet: area.sheet, r0: area.r0, c0: area.c0, r1, c1 };
  }
  return area.r1 - area.r0 + 1 === rows && area.c1 - area.c0 + 1 === cols ? area : null;
};

const numericAggregate = (values: readonly CellValue[], fn: MacDataFunction): CellValue => {
  const numbers = values.filter(isFiniteNumber).map((value) => value.value);
  if (fn === 'count') return { kind: 'number', value: numbers.length };
  if (numbers.length === 0) return { kind: 'blank' };
  switch (fn) {
    case 'sum':
      return { kind: 'number', value: numbers.reduce((sum, value) => sum + value, 0) };
    case 'average':
      return {
        kind: 'number',
        value: numbers.reduce((sum, value) => sum + value, 0) / numbers.length,
      };
    case 'min':
      return {
        kind: 'number',
        value: numbers.reduce(
          (minimum, value) => Math.min(minimum, value),
          Number.POSITIVE_INFINITY,
        ),
      };
    case 'max':
      return {
        kind: 'number',
        value: numbers.reduce(
          (maximum, value) => Math.max(maximum, value),
          Number.NEGATIVE_INFINITY,
        ),
      };
  }
};

/** Return true when a source range is covered by a table, pivot, or merge.
 * Subtotal intentionally refuses the whole construct because structural row
 * insertion can otherwise corrupt the object's definition. */
export function hasUnsupportedSubtotalObject(instance: SpreadsheetInstance, range: Range): boolean {
  const state = instance.store.getState();
  if (state.tables.tables.some((table) => rangesIntersect(table.range, range))) return true;
  for (const table of instance.workbook.getTables()) {
    const parsed = parseMacRange(instance.workbook, table.ref, table.sheetIndex);
    if (parsed && rangesIntersect(parsed.range, range)) return true;
  }
  for (const pivot of instance.workbook.getPivotTables()) {
    const pivotRange: Range = {
      sheet: pivot.sheetIndex,
      r0: pivot.top,
      c0: pivot.left,
      r1: pivot.top + Math.max(0, pivot.rows - 1),
      c1: pivot.left + Math.max(0, pivot.cols - 1),
    };
    if (rangesIntersect(pivotRange, range)) return true;
  }
  if (
    state.merges.byAnchor &&
    [...state.merges.byAnchor.values()].some((merge) => rangesIntersect(merge, range))
  ) {
    return true;
  }
  return instance.workbook.getMerges(range.sheet).some((merge) => rangesIntersect(merge, range));
}

export function supportsSubtotal(wb: WorkbookHandle): boolean {
  const names = wb.functionNames();
  if (names?.some((name) => name.toUpperCase() === 'SUBTOTAL')) return true;
  return wb.functionMetadata('SUBTOTAL') !== null;
}

const registerGoalSeekSolution = (
  request: GoalSeekRequest,
  sourceBytes: Uint8Array,
  solution: GoalSeekSolution,
): MacDataResult<GoalSeekSolution> => {
  goalSeekProofs.set(solution, {
    sourceBytes: new Uint8Array(sourceBytes),
    formulaCell: { ...request.formulaCell },
    changingCell: { ...request.changingCell },
    targetValue: request.targetValue,
    tolerance: request.tolerance,
    maxIterations: request.maxIterations,
    solution,
  });
  return success(solution);
};

/** Goal Seek runs entirely on a scratch workbook created from the latest
 * saved bytes. The live handle is only touched by the final controller
 * command after the caller has rechecked the source snapshots. */
export async function solveMacGoalSeek(
  instance: SpreadsheetInstance,
  request: GoalSeekRequest,
): Promise<MacDataResult<GoalSeekSolution>> {
  const { workbook: wb } = instance;
  if (!Number.isFinite(request.targetValue)) return failure('invalid', 'targetNotFinite');
  if (sameAddr(request.formulaCell, request.changingCell)) {
    return failure('invalid', 'sameCell');
  }
  const formula = wb.cellFormula(request.formulaCell);
  if (!formula) return failure('invalid', 'noFormula');
  if (wb.cellFormula(request.changingCell) !== null) {
    return failure('invalid', 'changingHasFormula');
  }
  const changing = wb.getValue(request.changingCell);
  if (!(changing.kind === 'blank' || isFiniteNumber(changing))) {
    return failure('invalid', 'changingNotNumeric');
  }
  const authorization = dataIntent(
    instance,
    'valueEdit',
    [request.changingCell],
    'mac.data.goalSeek',
  );
  if (!authorization.allowed) return failure('rejected', 'notEditable', authorization.reason);

  let scratch: WorkbookHandle | null = null;
  try {
    const sourceBytes = new Uint8Array(wb.save({ snapshot: true }));
    scratch = await WorkbookHandle.loadBytes(sourceBytes, { preferStub: wb.isStub });
    // A workbook in Manual calculation mode can carry stale cached results in
    // its saved bytes. Goal Seek always starts from a fresh calculation and
    // explicitly recalculates every trial for the same reason.
    scratch.recalc();
    if (scratch.cellFormula(request.formulaCell) !== formula) {
      return failure('unsupported', 'scratchUnsupported');
    }
    const initial = changing.kind === 'number' ? changing.value : 0;
    const tolerance = request.tolerance ?? GOAL_SEEK_MAX_CHANGE;
    if (!Number.isFinite(tolerance) || tolerance < 0) {
      return failure('invalid', 'invalidTolerance');
    }
    const requestedIterations = request.maxIterations ?? MAX_GOAL_SEEK_ITERATIONS;
    if (!Number.isFinite(requestedIterations) || requestedIterations < 1) {
      return failure('invalid', 'invalidIterations');
    }
    const maxIterations = Math.max(
      1,
      Math.min(MAX_GOAL_SEEK_ITERATIONS, Math.trunc(requestedIterations)),
    );
    // Returns the formula value for a trial changing value, or null when the
    // trial or the result is not a finite number.
    const evaluate = (value: number): number | null => {
      if (!Number.isFinite(value)) return null;
      scratch?.applyCellPatchAtomic([
        { addr: request.changingCell, value: { kind: 'number', value }, formula: null },
      ]);
      scratch?.recalc();
      return formulaNumber(scratch as WorkbookHandle, request.formulaCell);
    };
    const finish = (
      changingValue: number,
      formulaValue: number,
      iterations: number,
      converged: boolean,
    ): MacDataResult<GoalSeekSolution> =>
      registerGoalSeekSolution(request, sourceBytes, {
        changingValue,
        formulaValue,
        iterations,
        converged,
      });

    const initialValue = evaluate(initial);
    if (initialValue === null) return failure('invalid', 'nonFiniteResult');
    const target = request.targetValue;
    if (Math.abs(initialValue - target) < tolerance) return finish(initial, initialValue, 0, true);
    // Desktop Goal Seek is a secant search scaled by s = |x0| (1 when x0 is
    // 0): a first trial s/100 away, then secant steps capped at s/10 with the
    // cap doubling per step. A trial is kept when it lowers |residual| or
    // crosses the target; otherwise the half step is tried, and failing that
    // the search widens from the last kept point by the doubling cap.
    const scale = initial === 0 ? 1 : Math.abs(initial);
    const probe = GOAL_SEEK_PROBE * scale;
    const firstLimit = GOAL_SEEK_STEP_LIMIT * scale;
    const initialResidual = initialValue - target;
    const improves = (next: number | null, residual: number): next is number => {
      if (next === null) return false;
      const r = next - target;
      return Math.abs(r) < Math.abs(residual) || r > 0 !== residual > 0;
    };
    // The last finite trial is what the desktop app leaves in the cell.
    let lastX = initial;
    let lastValue = initialValue;
    const trial = (candidate: number): number | null => {
      const next = evaluate(candidate);
      if (next !== null) {
        lastX = candidate;
        lastValue = next;
      }
      return next;
    };
    const finishLast = (iterations: number): MacDataResult<GoalSeekSolution> =>
      finish(lastX, lastValue, iterations, false);

    let first: number | null = null;
    let firstValue = 0;
    let iterations = 1;
    for (const candidate of [initial + probe, initial - probe]) {
      const next = trial(candidate);
      if (improves(next, initialResidual)) {
        first = candidate;
        firstValue = next;
        break;
      }
    }
    for (let doublings = 1; first === null && doublings <= GOAL_SEEK_PROBE_DOUBLINGS; doublings++) {
      const candidate = initial - probe * 2 ** doublings;
      const next = trial(candidate);
      iterations += 1;
      if (improves(next, initialResidual)) {
        first = candidate;
        firstValue = next;
      }
    }
    if (first === null) return finishLast(iterations);
    if (Math.abs(firstValue - target) < tolerance)
      return finish(first, firstValue, iterations, true);

    const direction = first > initial ? 1 : -1;
    let previous = initial;
    let previousResidual = initialResidual;
    let x = first;
    let value = firstValue;
    let limit = firstLimit;
    // While every step is capped in the first trial's direction, positions are
    // measured from the first trial, as the desktop app does.
    let walking = true;
    let widening = 0;
    const steps = Math.min(GOAL_SEEK_SEARCH_STEPS, maxIterations - iterations);
    for (let step = 0; step < steps; step += 1) {
      iterations += 1;
      const residual = value - target;
      let candidate: number;
      if (widening !== 0) {
        candidate = x + widening * limit;
        limit *= 2;
        const next = trial(candidate);
        if (improves(next, residual)) {
          widening = 0;
          previous = x;
          previousResidual = residual;
          x = candidate;
          value = next;
          if (Math.abs(value - target) < tolerance) return finish(x, value, iterations, true);
        }
        continue;
      }
      let delta =
        residual === previousResidual
          ? -direction * Number.POSITIVE_INFINITY
          : -residual * ((x - previous) / (residual - previousResidual));
      const capped = Math.abs(delta) > limit;
      if (capped) delta = Math.sign(delta) * limit;
      if (capped && walking && Math.sign(delta) === direction) {
        candidate = first + direction * (2 * limit - firstLimit);
      } else {
        walking = false;
        candidate = x + delta;
      }
      limit *= 2;
      let next = trial(candidate);
      if (!improves(next, residual)) {
        walking = false;
        const half = x + (candidate - x) / 2;
        const halfValue = evaluate(half);
        if (!improves(halfValue, residual)) {
          widening = candidate > x ? 1 : -1;
          continue;
        }
        candidate = half;
        next = halfValue;
      }
      previous = x;
      previousResidual = residual;
      x = candidate;
      value = next;
      if (Math.abs(value - target) < tolerance) return finish(x, value, iterations, true);
    }
    return finishLast(iterations);
  } catch (error) {
    return failure('invalid', 'failed', thrownReason(error));
  } finally {
    scratch?.dispose();
  }
}

/** Apply a Goal Seek result (converged or the last iterate) exactly once through the interaction
 * controller. The controller rechecks protection/restrictions and creates the
 * real one-step undo/redo intent. */
export function commitMacGoalSeek(
  instance: SpreadsheetInstance,
  request: GoalSeekRequest,
  solution: GoalSeekSolution,
  formulaSnapshot: CellSnapshot,
  changingSnapshot: CellSnapshot,
): MacDataResult<void> {
  const currentFormula = cellSnapshot(instance.workbook, request.formulaCell);
  const currentChanging = cellSnapshot(instance.workbook, request.changingCell);
  if (
    !sameSnapshot(currentFormula, formulaSnapshot) ||
    !sameSnapshot(currentChanging, changingSnapshot)
  ) {
    return failure('stale', 'stale');
  }
  const proof = goalSeekProofs.get(solution);
  if (
    !proof ||
    !sameAddr(proof.formulaCell, request.formulaCell) ||
    !sameAddr(proof.changingCell, request.changingCell) ||
    !Object.is(proof.targetValue, request.targetValue) ||
    !Object.is(proof.tolerance, request.tolerance) ||
    !Object.is(proof.maxIterations, request.maxIterations) ||
    !Object.is(proof.solution.changingValue, solution.changingValue) ||
    !Object.is(proof.solution.formulaValue, solution.formulaValue) ||
    !Object.is(proof.solution.iterations, solution.iterations) ||
    proof.solution.converged !== solution.converged
  ) {
    return failure('stale', 'stale');
  }
  let currentBytes: Uint8Array;
  try {
    currentBytes = instance.workbook.save({ snapshot: true });
  } catch {
    return failure('stale', 'stale');
  }
  if (!sameSavedWorkbook(proof.sourceBytes, currentBytes)) {
    return failure('stale', 'stale');
  }
  const result = instance.commands.execute({
    type: 'cellBatch',
    operation: 'valueEdit',
    origin: 'ribbon',
    commandId: 'mac.data.goalSeek',
    changes: [
      {
        addr: request.changingCell,
        value: { kind: 'number', value: solution.changingValue },
        formula: null,
      },
    ],
  });
  return result.status === 'applied'
    ? success(undefined)
    : failure('rejected', 'writeFailed', result.rejections[0]?.reason);
}

/** Consolidate that matches columns by top-row label and/or rows by left-column label;
 *  an axis without labels is matched by position. Cells no source reaches stay empty. */
function planLabelledConsolidate(
  instance: SpreadsheetInstance,
  request: ConsolidateRequest,
  sources: readonly Range[],
  fallbackSheet: number,
): MacDataResult<ConsolidateSolution> {
  const wb = instance.workbook;
  const topRow = request.topRow === true;
  const leftColumn = request.leftColumn === true;
  const rowKeys = new Map<string, CellValue>();
  const colKeys = new Map<string, CellValue>();
  const buckets = new Map<string, CellValue[]>();
  const keyOf = (value: CellValue, position: number, labelled: boolean): string | null =>
    labelled ? (value.kind === 'blank' ? null : `t:${subtotalLabel(value)}`) : `p:${position}`;
  let sourceCells = 0;
  for (const range of sources) {
    const r0 = range.r0 + (topRow ? 1 : 0);
    const c0 = range.c0 + (leftColumn ? 1 : 0);
    sourceCells += (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);
    for (let row = r0; row <= range.r1; row += 1) {
      const rowLabel = leftColumn
        ? wb.getValue({ sheet: range.sheet, row, col: range.c0 })
        : ({ kind: 'blank' } as CellValue);
      const rk = keyOf(rowLabel, row - r0, leftColumn);
      if (rk === null) continue;
      if (!rowKeys.has(rk)) rowKeys.set(rk, rowLabel);
      for (let col = c0; col <= range.c1; col += 1) {
        const colLabel = topRow
          ? wb.getValue({ sheet: range.sheet, row: range.r0, col })
          : ({ kind: 'blank' } as CellValue);
        const ck = keyOf(colLabel, col - c0, topRow);
        if (ck === null) continue;
        if (!colKeys.has(ck)) colKeys.set(ck, colLabel);
        const bucketKey = `${rk}\u0000${ck}`;
        const list = buckets.get(bucketKey) ?? [];
        list.push(wb.getValue({ sheet: range.sheet, row, col }));
        buckets.set(bucketKey, list);
      }
    }
  }
  if (sourceCells > MAX_CONSOLIDATE_CELLS) return failure('unsupported', 'areaLimit');
  // Position keys stay in numeric order even when a later source is larger.
  const ordered = (keys: Map<string, CellValue>, labelled: boolean): [string, CellValue][] =>
    labelled ? [...keys] : [...keys].sort((a, b) => Number(a[0].slice(2)) - Number(b[0].slice(2)));
  const rowList = ordered(rowKeys, leftColumn);
  const colList = ordered(colKeys, topRow);
  const rows = rowList.length + (topRow ? 1 : 0);
  const cols = colList.length + (leftColumn ? 1 : 0);
  if (rowList.length === 0 || colList.length === 0) return failure('invalid', 'noSources');
  const destination = parseDestination(wb, request.destination, fallbackSheet, rows, cols);
  if (!destination) return failure('invalid', 'invalidDestination');
  if (sources.some((source) => rangesIntersect(source, destination))) {
    return failure('invalid', 'destinationOverlapsSource');
  }
  const output: CellValue[] = [];
  const blankValue: CellValue = { kind: 'blank' };
  if (topRow) {
    if (leftColumn) output.push(blankValue);
    for (const [, label] of colList) output.push(label);
  }
  for (const [rk, rowLabel] of rowList) {
    if (leftColumn) output.push(rowLabel);
    for (const [ck] of colList) {
      const values = buckets.get(`${rk}\u0000${ck}`);
      output.push(values ? numericAggregate(values, request.function) : blankValue);
    }
  }
  const authorization = dataIntent(
    instance,
    'valueEdit',
    cellAddresses(destination),
    'mac.data.consolidate',
  );
  if (!authorization.allowed) return failure('rejected', 'notEditable', authorization.reason);
  return success({ destination, values: output });
}

/** Position-based Consolidate. The destination must not overlap any source, so
 * a run can never overwrite its own inputs. Unsupported label/link/replace modes are intentionally absent from
 * this operation; the dialog exposes them as disabled affordances. */
export function planMacConsolidate(
  instance: SpreadsheetInstance,
  request: ConsolidateRequest,
): MacDataResult<ConsolidateSolution> {
  const wb = instance.workbook;
  const fallbackSheet = instance.store.getState().data.sheetIndex;
  if (request.sources.length === 0) return failure('invalid', 'noSources');
  if (!MAC_DATA_FUNCTIONS.has(request.function)) {
    return failure('invalid', 'invalidFunction');
  }
  const parsedSources: Range[] = [];
  for (const raw of request.sources) {
    const parsed = parseMacRange(wb, raw, fallbackSheet);
    if (!parsed) return failure('invalid', 'invalidSource');
    parsedSources.push(parsed.range);
  }
  const first = parsedSources[0];
  if (!first) return failure('invalid', 'noSources');
  if (request.topRow || request.leftColumn) {
    return planLabelledConsolidate(instance, request, parsedSources, fallbackSheet);
  }
  const rows = first.r1 - first.r0 + 1;
  const cols = first.c1 - first.c0 + 1;
  const area = rows * cols;
  if (
    area <= 0 ||
    area > MAX_CONSOLIDATE_CELLS ||
    area * parsedSources.length > MAX_CONSOLIDATE_CELLS
  ) {
    return failure('unsupported', 'areaLimit');
  }
  if (
    parsedSources.some(
      (range) => range.r1 - range.r0 + 1 !== rows || range.c1 - range.c0 + 1 !== cols,
    )
  ) {
    return failure('invalid', 'dimensionMismatch');
  }
  const destination = parseDestination(wb, request.destination, fallbackSheet, rows, cols);
  if (!destination) return failure('invalid', 'invalidDestination');
  if (parsedSources.some((source) => rangesIntersect(source, destination))) {
    return failure('invalid', 'destinationOverlapsSource');
  }
  const sourceValues = parsedSources.map((range) => {
    const values: CellValue[] = [];
    for (let row = range.r0; row <= range.r1; row += 1) {
      for (let col = range.c0; col <= range.c1; col += 1) {
        values.push(wb.getValue({ sheet: range.sheet, row, col }));
      }
    }
    return values;
  });
  const output: CellValue[] = [];
  for (let index = 0; index < area; index += 1) {
    output.push(
      numericAggregate(
        sourceValues.map((values) => values[index] ?? { kind: 'blank' }),
        request.function,
      ),
    );
  }
  const outputAddresses = cellAddresses(destination);
  const authorization = dataIntent(instance, 'valueEdit', outputAddresses, 'mac.data.consolidate');
  if (!authorization.allowed) return failure('rejected', 'notEditable', authorization.reason);
  return success({ destination, values: output });
}

export function commitMacConsolidate(
  instance: SpreadsheetInstance,
  _request: ConsolidateRequest,
  solution: ConsolidateSolution,
): MacDataResult<void> {
  const addresses = cellAddresses(solution.destination);
  const result = instance.commands.execute({
    type: 'cellBatch',
    operation: 'valueEdit',
    origin: 'ribbon',
    commandId: 'mac.data.consolidate',
    changes: addresses.map((addr, index) => ({
      addr,
      value: solution.values[index] ?? { kind: 'blank' },
      formula: null,
    })),
  });
  return result.status === 'applied'
    ? success(undefined)
    : failure('rejected', 'writeFailed', result.rejections[0]?.reason);
}

const subtotalFunctionCode = (fn: MacDataFunction): number => {
  switch (fn) {
    case 'average':
      return 1;
    case 'count':
      return 2;
    case 'max':
      return 4;
    case 'min':
      return 5;
    case 'sum':
      return 9;
  }
};

const subtotalLabel = (value: CellValue): string => {
  if (value.kind === 'blank') return '';
  if (value.kind === 'number') return String(value.value);
  if (value.kind === 'bool') return value.value ? 'TRUE' : 'FALSE';
  if (value.kind === 'text') return value.value;
  return value.text;
};

export function formatMacCellAddress(_wb: WorkbookHandle, addr: Addr): string {
  return formatA1Cell(addr.row, addr.col);
}

export function formatMacRangeAddress(wb: WorkbookHandle, range: Range): string {
  const sheet = wb.sheetName(range.sheet);
  const prefix = /[^A-Za-z0-9_]/.test(sheet) ? `'${sheet.replaceAll("'", "''")}'!` : `${sheet}!`;
  const first = formatMacCellAddress(wb, { sheet: range.sheet, row: range.r0, col: range.c0 });
  const last = formatMacCellAddress(wb, { sheet: range.sheet, row: range.r1, col: range.c1 });
  return `${prefix}${first === last ? first : `${first}:${last}`}`;
}

/** The relative same-sheet range a SUBTOTAL row sums, as Excel writes it. */
const formulaRange = (col: number, row0: number, row1: number): string =>
  `${colLetter(col)}${row0 + 1}:${colLetter(col)}${row1 + 1}`;

/** Final rows of each group's subtotal and of the grand total. Below: a group's
 *  row follows its data and the grand total comes last. Above: the grand total
 *  sits under the header and each group's row precedes its data. */
const subtotalRows = (
  range: Range,
  groups: readonly { readonly start: number; readonly insertAt: number }[],
  below: boolean,
): { groupRows: number[]; grandTotalRow: number } =>
  below
    ? {
        groupRows: groups.map((group, index) => group.insertAt + index),
        grandTotalRow: range.r1 + groups.length + 1,
      }
    : {
        groupRows: groups.map((group, index) => group.start + index + 1),
        grandTotalRow: range.r0 + 1,
      };

const subtotalOutputCells = (
  range: Range,
  columns: readonly number[],
  groups: readonly { readonly start: number; readonly insertAt: number }[],
  below = true,
): { outputCells: Addr[]; formulaCells: Addr[] } | null => {
  const cellsPerGroup = columns.length + 1;
  const maxGroupsByOutput = Math.floor(MAX_SUBTOTAL_OUTPUT_CELLS / cellsPerGroup);
  if (groups.length > MAX_SUBTOTAL_GROUPS || groups.length > maxGroupsByOutput) return null;
  const outputCells: Addr[] = [];
  const formulaCells: Addr[] = [];
  const addRow = (row: number): void => {
    outputCells.push({ sheet: range.sheet, row, col: range.c0 });
    for (const col of columns) {
      const addr = { sheet: range.sheet, row, col };
      outputCells.push(addr);
      formulaCells.push(addr);
    }
  };
  const { groupRows, grandTotalRow } = subtotalRows(range, groups, below);
  for (const row of [...groupRows, grandTotalRow]) {
    if (row >= MAX_ROW) return null;
    addRow(row);
  }
  return { outputCells, formulaCells };
};

/** Place the sheet's outline summary rows above their detail as one history entry. */
const recordOutlineSummaryAbove = (
  history: SpreadsheetInstance['history'],
  store: SpreadsheetStore,
  sheet: number,
): void => {
  mutators.setOutlineSummary(store, sheet, { below: false });
  history.push({
    undo: () => mutators.setOutlineSummary(store, sheet, { below: true }),
    redo: () => mutators.setOutlineSummary(store, sheet, { below: false }),
  });
};

/** Analyze and authorize a Subtotal operation before it starts any row
 * insertion. The returned group positions are in original worksheet rows. */
/** Subtotal rows are inserted across whole rows, so any data table they reach is refused. */
const blockSubtotalOverDataTable = (wb: WorkbookHandle, range: Range): boolean =>
  blockDataTableShift(wb, { sheet: range.sheet, r0: range.r0, c0: 0, r1: range.r1, c1: MAX_COL });

export function planMacSubtotal(
  instance: SpreadsheetInstance,
  request: SubtotalRequest,
): MacDataResult<SubtotalSolution> {
  const wb = instance.workbook;
  const state = instance.store.getState();
  if (!MAC_DATA_FUNCTIONS.has(request.function)) {
    return failure('invalid', 'invalidFunction');
  }
  const fallbackSheet = state.data.sheetIndex;
  const parsed = parseMacRange(wb, request.range, fallbackSheet);
  if (!parsed) return failure('invalid', 'invalidRange');
  const range = parsed.range;
  if (range.sheet !== state.data.sheetIndex) {
    return failure('unsupported', 'activeSheetOnly');
  }
  const rowCount = range.r1 - range.r0 + 1;
  if (rowCount < 2 || rowCount > MAX_SUBTOTAL_ROWS) return failure('invalid', 'needsDataRow');
  if (!supportsSubtotal(wb)) return failure('unsupported', 'engineUnsupported');
  if (hasUnsupportedSubtotalObject(instance, range))
    return failure('unsupported', 'unsupportedObjects');
  if (isSheetProtected(state, range.sheet)) return failure('rejected', 'protectedSheet');
  if (blockSubtotalOverDataTable(wb, range)) return failure('rejected', 'notEditable');
  const dataStart = range.r0 + 1;
  const groupColumn = request.groupByColumn;
  if (!Number.isInteger(groupColumn) || groupColumn < range.c0 || groupColumn > range.c1)
    return failure('invalid', 'groupOutsideRange');
  const columns = [...new Set(request.subtotalColumns)].filter(
    (col) => col >= range.c0 && col <= range.c1,
  );
  if (columns.length === 0) return failure('invalid', 'invalidColumn');
  if (columns.some((col) => !Number.isInteger(col))) return failure('invalid', 'invalidColumn');
  const maxGroupsByOutput = Math.floor(MAX_SUBTOTAL_OUTPUT_CELLS / (columns.length + 1));
  const maxGroups = Math.min(MAX_SUBTOTAL_GROUPS, maxGroupsByOutput);
  const groups: { start: number; end: number; insertAt: number; label: string }[] = [];
  const appendGroup = (group: {
    start: number;
    end: number;
    insertAt: number;
    label: string;
  }): boolean => {
    if (groups.length >= maxGroups) return false;
    groups.push(group);
    return true;
  };
  let currentStart = dataStart;
  let previousKey = subtotalLabel(
    wb.getValue({ sheet: range.sheet, row: dataStart, col: groupColumn }),
  );
  for (let row = dataStart + 1; row <= range.r1; row += 1) {
    const nextKey = subtotalLabel(wb.getValue({ sheet: range.sheet, row, col: groupColumn }));
    if (nextKey === previousKey) continue;
    if (!appendGroup({ start: currentStart, end: row - 1, insertAt: row, label: previousKey })) {
      return failure('unsupported', 'groupLimit');
    }
    currentStart = row;
    previousKey = nextKey;
  }
  if (
    !appendGroup({ start: currentStart, end: range.r1, insertAt: range.r1 + 1, label: previousKey })
  ) {
    return failure('unsupported', 'groupLimit');
  }
  if (groups.length === 0) return failure('invalid', 'noGroups');
  if (groups.some((group) => group.insertAt >= MAX_ROW)) {
    return failure('unsupported', 'noRoom');
  }

  // Future subtotal rows are shifted by insertions belonging to groups above.
  // Preflight every apply, undo, and redo intent at their eventual positions.
  const output = subtotalOutputCells(range, columns, groups, request.summaryBelow !== false);
  if (!output) return failure('unsupported', 'groupLimit');
  const intents = subtotalIntents(
    range,
    output.outputCells,
    output.formulaCells,
    request.pageBreaks === true,
  );
  const authorization = authorizeWrite(instance.store, intents, 'ribbon');
  if (!authorization.allowed) return failure('rejected', 'notEditable', authorization.reason);
  return success({ range: { ...range }, insertedRows: groups.length + 1, groups });
}

/** Commit Subtotal using the existing structural and outline history
 * primitives. All nested entries are held by one outer transaction; any
 * insertion, label, formula, or outline failure aborts that transaction. */
export function commitMacSubtotal(
  instance: SpreadsheetInstance,
  request: SubtotalRequest,
  solution: SubtotalSolution,
): MacDataResult<void> {
  const { store, workbook: wb, history } = instance;
  const parsed = parseMacRange(wb, request.range, store.getState().data.sheetIndex);
  if (!parsed) return failure('stale', 'stale');
  if (
    !solution.range ||
    !sameRange(parsed.range, solution.range) ||
    store.getState().data.sheetIndex !== solution.range.sheet
  ) {
    return failure('stale', 'stale');
  }
  if (solution.insertedRows !== solution.groups.length + 1) {
    return failure('stale', 'stale');
  }
  const range = { ...solution.range };
  if (blockSubtotalOverDataTable(wb, range)) return failure('rejected', 'notEditable');
  const labels = instance.i18n.strings.mac.data.subtotal;
  const functionCode = subtotalFunctionCode(request.function);
  const columns = [...new Set(request.subtotalColumns)].filter(
    (col) => col >= range.c0 && col <= range.c1,
  );
  if (columns.length === 0 || columns.some((col) => !Number.isInteger(col))) {
    return failure('stale', 'stale');
  }
  const below = request.summaryBelow !== false;
  const output = subtotalOutputCells(range, columns, solution.groups, below);
  if (!output) return failure('unsupported', 'groupLimit');
  const intents = subtotalIntents(
    range,
    output.outputCells,
    output.formulaCells,
    request.pageBreaks === true,
  );
  const write = (): void => {
    const writeTotalRow = (row: number, label: string, from: number, to: number): void => {
      wb.setText({ sheet: range.sheet, row, col: range.c0 }, label);
      wb.withBatchedRecalc(() => {
        for (const col of columns) {
          wb.setFormula(
            { sheet: range.sheet, row, col },
            `=SUBTOTAL(${functionCode},${formulaRange(col, from, to)})`,
          );
        }
      });
    };
    // Work bottom-up. Rows above the insertion point retain their original
    // indices; later upper insertions naturally shift formulas already made.
    for (let index = solution.groups.length - 1; index >= 0; index -= 1) {
      const group = solution.groups[index];
      if (!group) continue;
      const subtotalRow = below ? group.insertAt : group.start;
      if (!insertRows(store, wb, history, subtotalRow, 1))
        throw new Error('Could not insert subtotal row.');
      const shift = below ? 0 : 1;
      const label = labels.groupTotal.replace('{group}', group.label);
      writeTotalRow(subtotalRow, label, group.start + shift, group.end + shift);
      if (group.end >= group.start)
        groupRows(store, history, group.start + shift, group.end + shift, wb);
    }
    // The grand total sums every data row; SUBTOTAL skips the nested subtotals.
    // Excel's outline puts detail rows at level 3, subtotal rows at 2 and the
    // header and grand total at 1, so one more group spans data and subtotals.
    const dataStart = range.r0 + 1;
    const lastSubtotalRow = range.r1 + solution.groups.length;
    const grandTotalRow = below ? lastSubtotalRow + 1 : dataStart;
    if (!insertRows(store, wb, history, grandTotalRow, 1))
      throw new Error('Could not insert the grand total row.');
    if (below) {
      writeTotalRow(grandTotalRow, labels.grandTotal, dataStart, lastSubtotalRow - 1);
      groupRows(store, history, dataStart, lastSubtotalRow, wb);
    } else {
      // Above: the first group's subtotal row precedes the first data row.
      writeTotalRow(grandTotalRow, labels.grandTotal, dataStart + 2, lastSubtotalRow + 1);
      groupRows(store, history, dataStart + 1, lastSubtotalRow + 1, wb);
    }
    const finalRows = subtotalRows(range, solution.groups, below);
    if (!below) {
      // Summaries above: label cells are bold, the SUBTOTAL cells keep their
      // format, and the outline summarizes above its detail.
      recordFormatChange(history, store, () => {
        for (const row of [...finalRows.groupRows, finalRows.grandTotalRow]) {
          mutators.setCellFormat(store, { sheet: range.sheet, row, col: range.c0 }, { bold: true });
        }
      });
      if (outlineSummaryFor(store.getState().layout, range.sheet).below) {
        recordOutlineSummaryAbove(history, store, range.sheet);
      }
    }
    if (request.pageBreaks) {
      // Each group after the first starts below the subtotal rows inserted above
      // it; with summaries above, its page opens on its own subtotal row.
      const breakRows = below
        ? solution.groups.slice(1).map((group, index) => group.start + index + 1)
        : finalRows.groupRows.slice(1);
      if (breakRows.length > 0) {
        recordPageSetupChange(history, store, () => {
          const current = getPageSetup(store.getState(), range.sheet).manualPageBreakRows ?? [];
          mutators.setPageSetup(store, range.sheet, {
            manualPageBreakRows: [...new Set([...current, ...breakRows])].sort((a, b) => a - b),
          });
        });
      }
    }
  };
  try {
    const written = authorizedWrite({ store, history, intents, origin: 'ribbon', write });
    if (!written.ok) return failure('rejected', 'notEditable', written.decision.reason);
    return success(undefined);
  } catch (error) {
    return failure('invalid', 'failed', thrownReason(error));
  }
}

/** Rows of `range` below its header whose cells hold a SUBTOTAL formula. */
const subtotalRowsIn = (wb: WorkbookHandle, range: Range): number[] => {
  const rows: number[] = [];
  for (let row = range.r0 + 1; row <= range.r1; row += 1) {
    for (let col = range.c0; col <= range.c1; col += 1) {
      const formula = wb.cellFormula({ sheet: range.sheet, row, col });
      if (formula && /^=?\s*SUBTOTAL\(/i.test(formula)) {
        rows.push(row);
        break;
      }
    }
  }
  return rows;
};

/** Remove every subtotal row from the region and clear the outline it carried.
 *  Returns how many rows were removed; a region with none is a no-op. */
export function removeMacSubtotals(
  instance: SpreadsheetInstance,
  range: Range,
): MacDataResult<number> {
  const { store, workbook: wb, history } = instance;
  if (isSheetProtected(store.getState(), range.sheet)) return failure('rejected', 'protectedSheet');
  if (blockSubtotalOverDataTable(wb, range)) return failure('rejected', 'notEditable');
  const rows = subtotalRowsIn(wb, range);
  if (rows.length === 0) return success(0);
  try {
    const written = authorizedWrite({
      store,
      history,
      intents: subtotalRemovalIntents(range),
      origin: 'ribbon',
      write: () => {
        // Bottom-up, so the rows still to delete keep their indices.
        for (const row of [...rows].reverse()) {
          if (!deleteRows(store, wb, history, row, 1))
            throw new Error('Could not delete subtotal row.');
        }
        const lastRow = range.r1 - rows.length;
        // Subtotal outlines nest up to three levels deep.
        for (let level = 0; level < 3; level += 1)
          ungroupRows(store, history, range.r0, lastRow, wb);
      },
    });
    if (!written.ok) return failure('rejected', 'notEditable', written.decision.reason);
    return success(rows.length);
  } catch (error) {
    return failure('invalid', 'failed', thrownReason(error));
  }
}

/** Replace the region's existing subtotals with a new set as one history step.
 *  `makeRequest` receives the region as it stands once the old rows are gone. */
export function replaceMacSubtotals(
  instance: SpreadsheetInstance,
  region: Range,
  makeRequest: (target: Range) => SubtotalRequest,
): MacDataResult<void> {
  const { store, history } = instance;
  // A step that fails rolls the whole replacement back, carrying its own result out.
  class Failed extends Error {
    constructor(readonly result: MacDataResult<void>) {
      super('subtotal replacement failed');
    }
  }
  const pageBreaks = makeRequest(region).pageBreaks === true;
  try {
    const written = authorizedWrite({
      store,
      history,
      intents: subtotalReplaceIntents(region, pageBreaks),
      origin: 'ribbon',
      write: () => {
        const removed = removeMacSubtotals(instance, region);
        if (!removed.ok) throw new Failed(removed);
        const request = makeRequest({ ...region, r1: region.r1 - removed.value });
        const plan = planMacSubtotal(instance, request);
        if (!plan.ok) throw new Failed(plan);
        const committed = commitMacSubtotal(instance, request, plan.value);
        if (!committed.ok) throw new Failed(committed);
      },
    });
    if (!written.ok) return failure('rejected', 'notEditable', written.decision.reason);
    return success(undefined);
  } catch (error) {
    if (error instanceof Failed) return error.result;
    return failure('invalid', 'failed', thrownReason(error));
  }
}

export { addrRange, cellPatchFromSnapshot, cellSnapshot, sameSnapshot };
