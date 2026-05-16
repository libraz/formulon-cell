import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore, State } from '../store/store.js';
import type { History } from './history.js';
import { isCellWritable, warnProtected } from './protection.js';

export interface DefinedNameEntry {
  name: string;
  formula: string;
}

export type DefinedNameMutationResult =
  | { ok: true; entry: DefinedNameEntry }
  | { ok: false; reason: 'empty-name' | 'empty-formula' | 'unsupported' | 'engine-failed' };

export type DefinedNameDeleteResult =
  | { ok: true; entry: DefinedNameEntry }
  | { ok: false; reason: 'empty-name' | 'unsupported' | 'engine-failed' };

export type CreateDefinedNamesSource = 'top-row' | 'bottom-row' | 'left-column' | 'right-column';

export interface CreateDefinedNamesResult {
  ok: true;
  entries: DefinedNameEntry[];
}

type DefinedNamesSnapshot = Map<string, string>;

const colLetter = (n: number): string => {
  let v = n;
  let out = '';
  do {
    out = String.fromCharCode(65 + (v % 26)) + out;
    v = Math.floor(v / 26) - 1;
  } while (v >= 0);
  return out;
};

const absRef = (row: number, col: number): string => `$${colLetter(col)}$${row + 1}`;

const absRangeRef = (r0: number, c0: number, r1: number, c1: number): string =>
  `${absRef(r0, c0)}:${absRef(r1, c1)}`;

const cellText = (state: State, sheet: number, row: number, col: number): string => {
  const cell = state.data.cells.get(`${sheet}:${row}:${col}`);
  if (!cell) return '';
  if (cell.value.kind === 'text') return cell.value.value;
  if (cell.value.kind === 'number') return String(cell.value.value);
  return '';
};

const sanitizeName = (raw: string, fallback: string): string => {
  const compact = raw.trim().replace(/\s+/g, '_').replace(/[^A-Za-z0-9_.\\]/g, '_');
  const base = compact || fallback;
  return /^[A-Za-z_\\]/.test(base) ? base : `_${base}`;
};

const uniqueName = (base: string, used: Set<string>): string => {
  let candidate = base;
  let suffix = 2;
  while (used.has(candidate.toLowerCase())) {
    candidate = `${base}_${suffix}`;
    suffix += 1;
  }
  used.add(candidate.toLowerCase());
  return candidate;
};

const captureDefinedNamesSnapshot = (wb: WorkbookHandle): DefinedNamesSnapshot =>
  new Map([...wb.definedNames()].map((entry) => [entry.name, entry.formula]));

const sameDefinedNamesSnapshot = (
  a: DefinedNamesSnapshot,
  b: DefinedNamesSnapshot,
): boolean => {
  if (a.size !== b.size) return false;
  for (const [name, formula] of a) {
    if (b.get(name) !== formula) return false;
  }
  return true;
};

const applyDefinedNamesSnapshot = (wb: WorkbookHandle, snap: DefinedNamesSnapshot): void => {
  for (const entry of [...wb.definedNames()]) {
    if (!snap.has(entry.name)) wb.setDefinedNameEntry(entry.name, '');
  }
  for (const [name, formula] of snap) wb.setDefinedNameEntry(name, formula);
};

export function recordDefinedNamesChange<T>(
  history: History | null,
  wb: WorkbookHandle,
  mutate: () => T,
): T {
  const before = captureDefinedNamesSnapshot(wb);
  const result = mutate();
  const after = captureDefinedNamesSnapshot(wb);
  if (history && !history.isReplaying() && !sameDefinedNamesSnapshot(before, after)) {
    history.push({
      undo: () => applyDefinedNamesSnapshot(wb, before),
      redo: () => applyDefinedNamesSnapshot(wb, after),
    });
  }
  return result;
}

/** Snapshot workbook-scoped defined names into a stable array. */
export function listDefinedNames(wb: WorkbookHandle): DefinedNameEntry[] {
  return [...wb.definedNames()];
}

/** Add or replace a workbook-scoped defined name. */
export function upsertDefinedName(
  wb: WorkbookHandle,
  name: string,
  formula: string,
): DefinedNameMutationResult {
  const trimmedName = name.trim();
  const trimmedFormula = formula.trim();
  if (!trimmedName) return { ok: false, reason: 'empty-name' };
  if (!trimmedFormula) return { ok: false, reason: 'empty-formula' };
  if (!wb.capabilities.definedNameMutate) return { ok: false, reason: 'unsupported' };
  if (!wb.setDefinedNameEntry(trimmedName, trimmedFormula)) {
    return { ok: false, reason: 'engine-failed' };
  }
  return { ok: true, entry: { name: trimmedName, formula: trimmedFormula } };
}

/** Remove a workbook-scoped defined name. */
export function deleteDefinedName(wb: WorkbookHandle, name: string): DefinedNameDeleteResult {
  const trimmedName = name.trim();
  if (!trimmedName) return { ok: false, reason: 'empty-name' };
  if (!wb.capabilities.definedNameMutate) return { ok: false, reason: 'unsupported' };
  if (!wb.setDefinedNameEntry(trimmedName, '')) return { ok: false, reason: 'engine-failed' };
  return { ok: true, entry: { name: trimmedName, formula: '' } };
}

export function createDefinedNamesFromSelection(
  state: State,
  wb: WorkbookHandle,
  source: CreateDefinedNamesSource,
): CreateDefinedNamesResult | { ok: false; reason: 'unsupported' | 'empty-selection' | 'engine-failed' } {
  if (!wb.capabilities.definedNameMutate) return { ok: false, reason: 'unsupported' };
  const sheet = state.selection.range.sheet;
  const r = state.selection.range;
  const used = new Set([...wb.definedNames()].map((entry) => entry.name.toLowerCase()));
  const entries: DefinedNameEntry[] = [];

  if (source === 'top-row') {
    if (r.r0 >= r.r1) return { ok: false, reason: 'empty-selection' };
    for (let col = r.c0; col <= r.c1; col += 1) {
      const label = cellText(state, sheet, r.r0, col);
      const name = uniqueName(sanitizeName(label, `Column_${colLetter(col)}`), used);
      const formula = `=${absRangeRef(r.r0 + 1, col, r.r1, col)}`;
      const result = upsertDefinedName(wb, name, formula);
      if (!result.ok) return { ok: false, reason: 'engine-failed' };
      entries.push(result.entry);
    }
    return { ok: true, entries };
  }

  if (source === 'bottom-row') {
    if (r.r0 >= r.r1) return { ok: false, reason: 'empty-selection' };
    for (let col = r.c0; col <= r.c1; col += 1) {
      const label = cellText(state, sheet, r.r1, col);
      const name = uniqueName(sanitizeName(label, `Column_${colLetter(col)}`), used);
      const formula = `=${absRangeRef(r.r0, col, r.r1 - 1, col)}`;
      const result = upsertDefinedName(wb, name, formula);
      if (!result.ok) return { ok: false, reason: 'engine-failed' };
      entries.push(result.entry);
    }
    return { ok: true, entries };
  }

  if (source === 'left-column') {
    if (r.c0 >= r.c1) return { ok: false, reason: 'empty-selection' };
    for (let row = r.r0; row <= r.r1; row += 1) {
      const label = cellText(state, sheet, row, r.c0);
      const name = uniqueName(sanitizeName(label, `Row_${row + 1}`), used);
      const formula = `=${absRangeRef(row, r.c0 + 1, row, r.c1)}`;
      const result = upsertDefinedName(wb, name, formula);
      if (!result.ok) return { ok: false, reason: 'engine-failed' };
      entries.push(result.entry);
    }
    return { ok: true, entries };
  }

  if (r.c0 >= r.c1) return { ok: false, reason: 'empty-selection' };
  for (let row = r.r0; row <= r.r1; row += 1) {
    const label = cellText(state, sheet, row, r.c1);
    const name = uniqueName(sanitizeName(label, `Row_${row + 1}`), used);
    const formula = `=${absRangeRef(row, r.c0, row, r.c1 - 1)}`;
    const result = upsertDefinedName(wb, name, formula);
    if (!result.ok) return { ok: false, reason: 'engine-failed' };
    entries.push(result.entry);
  }
  return { ok: true, entries };
}

export function insertDefinedNameFormula(
  state: State,
  wb: WorkbookHandle,
  name: string,
  store?: SpreadsheetStore,
): { addr: Addr; formula: string } | null {
  const trimmed = name.trim();
  if (!trimmed) return null;
  const addr = state.selection.active;
  if (store && !isCellWritable(store.getState(), addr)) {
    warnProtected(addr);
    return null;
  }
  const formula = `=${trimmed}`;
  wb.setFormula(addr, formula);
  return { addr, formula };
}
