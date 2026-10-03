import { parseRangeRef } from '../engine/range-resolver.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import { type CellSnapshot, WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import type { SpreadsheetInstance } from '../mount/types.js';
import type { OperationIntent, PermissionDecision } from './interaction-policy.js';
import { groupRows } from './outline.js';
import { isSheetProtected } from './protection.js';
import { insertRows } from './structure.js';

/** Data commands exposed by the Mac ribbon. The limits are
 * deliberately finite because these commands materialize every output cell
 * before committing it. */
export const MAX_CONSOLIDATE_CELLS = 100_000;
export const MAX_SUBTOTAL_ROWS = 100_000;
export const MAX_SUBTOTAL_OUTPUT_CELLS = 100_000;
export const MAX_SUBTOTAL_GROUPS = 1_000;
export const MAX_GOAL_SEEK_ITERATIONS = 100;

export type MacDataFunction = 'sum' | 'average' | 'count' | 'min' | 'max';

const MAC_DATA_FUNCTIONS: ReadonlySet<MacDataFunction> = new Set([
  'sum',
  'average',
  'count',
  'min',
  'max',
]);

/** Keys of the localized `macData.errors` table. */
export type MacDataErrorCode = keyof Strings['macData']['errors'];

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
  error.reason ?? strings.macData.errors[error.code];

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

const MAX_ROW = 1_048_575;
const MAX_COL = 16_383;

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

const sameAddr = (a: Addr, b: Addr): boolean =>
  a.sheet === b.sheet && a.row === b.row && a.col === b.col;

const sameRange = (a: Range, b: Range): boolean =>
  a.sheet === b.sheet && a.r0 === b.r0 && a.c0 === b.c0 && a.r1 === b.r1 && a.c1 === b.c1;

const sameBytes = (a: Uint8Array, b: Uint8Array): boolean =>
  a.length === b.length && a.every((byte, index) => byte === b[index]);

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
    const wanted = parsed.sheetName.toLowerCase();
    sheet = -1;
    for (let i = 0; i < wb.sheetCount; i += 1) {
      if (wb.sheetName(i).toLowerCase() === wanted) {
        sheet = i;
        break;
      }
    }
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

export function rangeArea(range: Range): number {
  return (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);
}

const rangesOverlap = (a: Range, b: Range): boolean =>
  a.sheet === b.sheet && a.r0 <= b.r1 && a.r1 >= b.r0 && a.c0 <= b.c1 && a.c1 >= b.c0;

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

interface SubtotalIntents {
  readonly apply: readonly OperationIntent[];
  readonly undo: readonly OperationIntent[];
  readonly redo: readonly OperationIntent[];
}

const subtotalIntents = (
  range: Range,
  outputCells: readonly Addr[],
  formulaCells: readonly Addr[],
): SubtotalIntents => {
  const cellsEffect = { kind: 'cells' as const, cells: outputCells };
  const formulasEffect = {
    kind: 'cells' as const,
    cells: formulaCells,
    includesFormula: true,
  };
  const formatEffect = { kind: 'workbook' as const };
  const structuralEffects = [
    formatEffect,
    { kind: 'range' as const, range: { ...range } },
  ] as const;
  const make = (
    origin: 'ribbon' | 'undo' | 'redo',
    structuralOperation: 'insertRows' | 'deleteRows',
  ): readonly OperationIntent[] => [
    {
      operation: structuralOperation,
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
  ];
  return {
    apply: make('ribbon', 'insertRows'),
    undo: make('undo', 'deleteRows'),
    redo: make('redo', 'insertRows'),
  };
};

const firstDeniedSubtotalIntent = (
  instance: SpreadsheetInstance,
  intents: SubtotalIntents,
): PermissionDecision => {
  let firstDenied: PermissionDecision | null = null;
  for (const intent of [...intents.apply, ...intents.undo, ...intents.redo]) {
    const decision = instance.commands.canExecute(intent);
    if (!decision.allowed && firstDenied === null) firstDenied = decision;
  }
  return firstDenied ?? { allowed: true };
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
  if (state.tables.tables.some((table) => rangesOverlap(table.range, range))) return true;
  for (const table of instance.workbook.getTables()) {
    const parsed = parseMacRange(instance.workbook, table.ref, table.sheetIndex);
    if (parsed && rangesOverlap(parsed.range, range)) return true;
  }
  for (const pivot of instance.workbook.getPivotTables()) {
    const pivotRange: Range = {
      sheet: pivot.sheetIndex,
      r0: pivot.top,
      c0: pivot.left,
      r1: pivot.top + Math.max(0, pivot.rows - 1),
      c1: pivot.left + Math.max(0, pivot.cols - 1),
    };
    if (rangesOverlap(pivotRange, range)) return true;
  }
  if (
    state.merges.byAnchor &&
    [...state.merges.byAnchor.values()].some((merge) => rangesOverlap(merge, range))
  ) {
    return true;
  }
  return instance.workbook.getMerges(range.sheet).some((merge) => rangesOverlap(merge, range));
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
  let sourceBytes: Uint8Array | null = null;
  try {
    sourceBytes = new Uint8Array(wb.save());
    scratch = await WorkbookHandle.loadBytes(sourceBytes, { preferStub: wb.isStub });
    // A workbook in Manual calculation mode can carry stale cached results in
    // its saved bytes. Goal Seek always starts from a fresh calculation and
    // explicitly recalculates every trial for the same reason.
    scratch.recalc();
    if (scratch.cellFormula(request.formulaCell) !== formula) {
      return failure('unsupported', 'scratchUnsupported');
    }
    const initial = changing.kind === 'number' ? changing.value : 0;
    const requestedTolerance = request.tolerance ?? 1e-8;
    if (!Number.isFinite(requestedTolerance) || requestedTolerance < 0) {
      return failure('invalid', 'invalidTolerance');
    }
    const tolerance = Math.max(
      Number.EPSILON * Math.max(1, Math.abs(request.targetValue)),
      requestedTolerance,
    );
    const requestedIterations = request.maxIterations ?? MAX_GOAL_SEEK_ITERATIONS;
    if (!Number.isFinite(requestedIterations) || requestedIterations < 1) {
      return failure('invalid', 'invalidIterations');
    }
    const maxIterations = Math.max(
      1,
      Math.min(MAX_GOAL_SEEK_ITERATIONS, Math.trunc(requestedIterations)),
    );
    const evaluate = (value: number): number | null => {
      if (!Number.isFinite(value)) return null;
      scratch?.applyCellPatchAtomic([
        { addr: request.changingCell, value: { kind: 'number', value }, formula: null },
      ]);
      scratch?.recalc();
      const formulaValue = formulaNumber(scratch as WorkbookHandle, request.formulaCell);
      return formulaValue === null ? null : formulaValue - request.targetValue;
    };
    const f0 = evaluate(initial);
    if (f0 === null) return failure('invalid', 'nonFiniteResult');
    if (Math.abs(f0) <= tolerance) {
      return registerGoalSeekSolution(request, sourceBytes, {
        changingValue: initial,
        formulaValue: request.targetValue + f0,
        iterations: 0,
      });
    }
    let x0 = initial;
    let fPrev = f0;
    let x1 = initial === 0 ? 1 : initial + Math.max(1, Math.abs(initial) * 0.05);
    if (x1 === x0) x1 = x0 + 1;
    let f1 = evaluate(x1);
    if (f1 === null) return failure('invalid', 'nonFiniteResult');
    if (Math.abs(f1) <= tolerance) {
      return registerGoalSeekSolution(request, sourceBytes, {
        changingValue: x1,
        formulaValue: request.targetValue + f1,
        iterations: 1,
      });
    }

    let bracketLow: number | null = null;
    let bracketHigh: number | null = null;
    let bracketLowValue = 0;
    if (f0 * f1 < 0) {
      bracketLow = Math.min(x0, x1);
      bracketHigh = Math.max(x0, x1);
      if (x0 < x1) {
        bracketLowValue = f0;
      } else {
        bracketLowValue = f1;
      }
    }

    for (let iteration = 2; iteration <= maxIterations; iteration += 1) {
      let candidate: number;
      const denominator = f1 - fPrev;
      if (Math.abs(denominator) <= Number.EPSILON * Math.max(1, Math.abs(f1), Math.abs(fPrev))) {
        candidate =
          bracketLow !== null && bracketHigh !== null ? (bracketLow + bracketHigh) / 2 : x1 + 1;
      } else {
        candidate = x1 - (f1 * (x1 - x0)) / denominator;
      }
      if (!Number.isFinite(candidate)) candidate = x1 + Math.max(1, Math.abs(x1) * 0.1);
      if (
        bracketLow !== null &&
        bracketHigh !== null &&
        (candidate <= bracketLow || candidate >= bracketHigh)
      ) {
        candidate = (bracketLow + bracketHigh) / 2;
      }
      const nextResidual = evaluate(candidate);
      if (nextResidual === null) return failure('invalid', 'nonFiniteResult');
      if (Math.abs(nextResidual) <= tolerance) {
        return registerGoalSeekSolution(request, sourceBytes, {
          changingValue: candidate,
          formulaValue: request.targetValue + nextResidual,
          iterations: iteration,
        });
      }
      if (bracketLow !== null && bracketHigh !== null) {
        if (bracketLowValue * nextResidual <= 0) {
          bracketHigh = candidate;
        } else {
          bracketLow = candidate;
          bracketLowValue = nextResidual;
        }
      } else if (f1 * nextResidual < 0) {
        bracketLow = Math.min(x1, candidate);
        bracketHigh = Math.max(x1, candidate);
        if (x1 < candidate) {
          bracketLowValue = f1;
        } else {
          bracketLowValue = nextResidual;
        }
      }
      x0 = x1;
      fPrev = f1;
      x1 = candidate;
      f1 = nextResidual;
    }
    return failure('nonconverged', 'notConverged');
  } catch (error) {
    return failure('invalid', 'failed', thrownReason(error));
  } finally {
    scratch?.dispose();
  }
}

/** Apply a converged Goal Seek value exactly once through the interaction
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
    !Object.is(proof.solution.iterations, solution.iterations)
  ) {
    return failure('stale', 'stale');
  }
  let currentBytes: Uint8Array;
  try {
    currentBytes = instance.workbook.save();
  } catch {
    return failure('stale', 'stale');
  }
  if (!sameBytes(proof.sourceBytes, currentBytes)) {
    return failure('stale', 'stale');
  }
  const result = instance.commands.execute({
    type: 'cellBatch',
    operation: 'valueEdit',
    origin: 'ribbon',
    commandId: 'mac.data.goalSeek',
    denied: 'reject',
    changes: [
      {
        addr: request.changingCell,
        value: { kind: 'number', value: solution.changingValue },
        formula: null,
      },
    ],
  });
  return result.status === 'applied' || (result.status === 'noop' && result.rejected.length === 0)
    ? success(undefined)
    : failure('rejected', 'writeFailed', result.rejected[0]?.reason);
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
  if (parsedSources.some((source) => rangesOverlap(source, destination))) {
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
    denied: 'reject',
    changes: addresses.map((addr, index) => ({
      addr,
      value: solution.values[index] ?? { kind: 'blank' },
      formula: null,
    })),
  });
  return result.status === 'applied' || (result.status === 'noop' && result.rejected.length === 0)
    ? success(undefined)
    : failure('rejected', 'writeFailed', result.rejected[0]?.reason);
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

const columnName = (col: number): string => {
  let n = col + 1;
  let text = '';
  while (n > 0) {
    const digit = (n - 1) % 26;
    text = String.fromCharCode(65 + digit) + text;
    n = Math.floor((n - 1) / 26);
  }
  return text;
};

export function formatMacCellAddress(_wb: WorkbookHandle, addr: Addr): string {
  return `${columnName(addr.col)}${addr.row + 1}`;
}

export function formatMacRangeAddress(wb: WorkbookHandle, range: Range): string {
  const sheet = wb.sheetName(range.sheet);
  const prefix = /[^A-Za-z0-9_]/.test(sheet) ? `'${sheet.replaceAll("'", "''")}'!` : `${sheet}!`;
  const first = formatMacCellAddress(wb, { sheet: range.sheet, row: range.r0, col: range.c0 });
  const last = formatMacCellAddress(wb, { sheet: range.sheet, row: range.r1, col: range.c1 });
  return `${prefix}${first === last ? first : `${first}:${last}`}`;
}

const formulaRange = (sheet: string, col: number, row0: number, row1: number): string => {
  const quoted = /[^A-Za-z0-9_]/.test(sheet) ? `'${sheet.replaceAll("'", "''")}'` : sheet;
  return `${quoted}!$${columnName(col)}$${row0 + 1}:$${columnName(col)}$${row1 + 1}`;
};

const subtotalOutputCells = (
  range: Range,
  columns: readonly number[],
  groups: readonly { readonly insertAt: number }[],
): { outputCells: Addr[]; formulaCells: Addr[] } | null => {
  const cellsPerGroup = columns.length + 1;
  const maxGroupsByOutput = Math.floor(MAX_SUBTOTAL_OUTPUT_CELLS / cellsPerGroup);
  if (groups.length > MAX_SUBTOTAL_GROUPS || groups.length > maxGroupsByOutput) return null;
  const outputCells: Addr[] = [];
  const formulaCells: Addr[] = [];
  for (const [index, group] of groups.entries()) {
    const finalSubtotalRow = group.insertAt + index;
    if (finalSubtotalRow >= MAX_ROW) return null;
    outputCells.push({ sheet: range.sheet, row: finalSubtotalRow, col: range.c0 });
    for (const col of columns) {
      const addr = { sheet: range.sheet, row: finalSubtotalRow, col };
      outputCells.push(addr);
      formulaCells.push(addr);
    }
  }
  return { outputCells, formulaCells };
};

/** Analyze and authorize a Subtotal operation before it starts any row
 * insertion. The returned group positions are in original worksheet rows. */
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
  const output = subtotalOutputCells(range, columns, groups);
  if (!output) return failure('unsupported', 'groupLimit');
  const intents = subtotalIntents(range, output.outputCells, output.formulaCells);
  const authorization = firstDeniedSubtotalIntent(instance, intents);
  if (!authorization.allowed) return failure('rejected', 'notEditable', authorization.reason);
  return success({ range: { ...range }, insertedRows: groups.length, groups });
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
  if (solution.insertedRows !== solution.groups.length) {
    return failure('stale', 'stale');
  }
  const range = { ...solution.range };
  const sheetName = wb.sheetName(range.sheet);
  const functionCode = subtotalFunctionCode(request.function);
  const columns = [...new Set(request.subtotalColumns)].filter(
    (col) => col >= range.c0 && col <= range.c1,
  );
  if (columns.length === 0 || columns.some((col) => !Number.isInteger(col))) {
    return failure('stale', 'stale');
  }
  const output = subtotalOutputCells(range, columns, solution.groups);
  if (!output) return failure('unsupported', 'groupLimit');
  const intents = subtotalIntents(range, output.outputCells, output.formulaCells);
  const authorization = firstDeniedSubtotalIntent(instance, intents);
  if (!authorization.allowed) return failure('rejected', 'notEditable', authorization.reason);
  let token: ReturnType<typeof history.begin>;
  try {
    token = history.begin({
      replayAuthorization: { undo: intents.undo, redo: intents.redo },
    });
  } catch (error) {
    return failure('invalid', 'failed', thrownReason(error));
  }
  try {
    // Work bottom-up. Rows above the insertion point retain their original
    // indices; later upper insertions naturally shift formulas already made.
    for (let index = solution.groups.length - 1; index >= 0; index -= 1) {
      const group = solution.groups[index];
      if (!group) continue;
      if (!insertRows(store, wb, history, group.insertAt, 1))
        throw new Error('Could not insert subtotal row.');
      const subtotalRow = group.insertAt;
      wb.setText({ sheet: range.sheet, row: subtotalRow, col: range.c0 }, `${group.label} Total`);
      wb.withBatchedRecalc(() => {
        for (const col of columns) {
          wb.setFormula(
            { sheet: range.sheet, row: subtotalRow, col },
            `=SUBTOTAL(${functionCode},${formulaRange(sheetName, col, group.start, group.end)})`,
          );
        }
      });
      if (group.end >= group.start) groupRows(store, history, group.start, group.end, wb);
    }
    history.end(token);
    return success(undefined);
  } catch (error) {
    try {
      history.abort(token);
    } catch (abortError) {
      return failure('invalid', 'failed', thrownReason(abortError));
    }
    return failure('invalid', 'failed', thrownReason(error));
  }
}

export { addrRange, cellPatchFromSnapshot, cellSnapshot, sameSnapshot };
