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
  | {
      ok: false;
      reason: 'empty-name' | 'invalid-name' | 'empty-formula' | 'unsupported' | 'engine-failed';
    };

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

const colIndexFromLetters = (letters: string): number => {
  let col = 0;
  for (let i = 0; i < letters.length; i += 1) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return col;
};

/**
 * Whether `name` collides with a cell reference and so may not be used as a
 * defined name. Covers A1-style addresses within the sheet grid (up to column
 * XFD / row 1048576) and R1C1-style tokens (R, C, RC, R1C1, …).
 */
const looksLikeCellRef = (name: string): boolean => {
  const up = name.toUpperCase();
  const a1 = up.match(/^([A-Z]{1,3})([0-9]{1,7})$/);
  if (a1) {
    const col = colIndexFromLetters(a1[1] ?? '');
    const row = Number.parseInt(a1[2] ?? '', 10);
    if (col >= 1 && col <= 16384 && row >= 1 && row <= 1048576) return true;
  }
  if (/^R[0-9]*C[0-9]*$/.test(up)) return true;
  return false;
};

/**
 * Validate a workbook-scoped defined name against spreadsheet naming rules:
 * the first character must be a letter, underscore, or backslash; the rest may
 * add digits and periods; and the name may not be a cell reference (A1 or R1C1
 * style), a bare `R`/`C`, contain spaces/illegal characters, or exceed 255
 * characters.
 */
export function isValidDefinedName(raw: string): boolean {
  const name = raw.trim();
  if (!name || name.length > 255) return false;
  if (!/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(name)) return false;
  // Bare R/C are reserved for row/column shorthand in R1C1 mode.
  if (/^[RC]$/i.test(name)) return false;
  if (looksLikeCellRef(name)) return false;
  return true;
}

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
  const compact = raw
    .trim()
    .replace(/\s+/g, '_')
    .replace(/[^A-Za-z0-9_.\\]/g, '_');
  const base = compact || fallback;
  const prefixed = /^[A-Za-z_\\]/.test(base) ? base : `_${base}`;
  // A label like "A1" sanitizes to a valid identifier but collides with a cell
  // reference, which the engine (and spreadsheets) reject as a name — prefix it
  // so "Create from Selection" always yields a usable name.
  return looksLikeCellRef(prefixed) ? `_${prefixed}` : prefixed;
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

const sameDefinedNamesSnapshot = (a: DefinedNamesSnapshot, b: DefinedNamesSnapshot): boolean => {
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
  // Undo/redo of a name change must recompute dependents too (H-38).
  wb.recalc();
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
  if (!isValidDefinedName(trimmedName)) return { ok: false, reason: 'invalid-name' };
  if (!trimmedFormula) return { ok: false, reason: 'empty-formula' };
  if (!wb.capabilities.definedNameMutate) return { ok: false, reason: 'unsupported' };
  if (!wb.setDefinedNameEntry(trimmedName, trimmedFormula)) {
    return { ok: false, reason: 'engine-failed' };
  }
  // Defining/redefining a name changes what `=MyRange` resolves to — recompute
  // so dependent cells and a subsequent save reflect the new target (H-38).
  wb.recalc();
  return { ok: true, entry: { name: trimmedName, formula: trimmedFormula } };
}

/** Remove a workbook-scoped defined name. */
export function deleteDefinedName(wb: WorkbookHandle, name: string): DefinedNameDeleteResult {
  const trimmedName = name.trim();
  if (!trimmedName) return { ok: false, reason: 'empty-name' };
  if (!wb.capabilities.definedNameMutate) return { ok: false, reason: 'unsupported' };
  if (!wb.setDefinedNameEntry(trimmedName, '')) return { ok: false, reason: 'engine-failed' };
  // Removing a name usually turns `=MyRange` into #NAME? — recompute so the
  // change is reflected in dependents and on save (H-38).
  wb.recalc();
  return { ok: true, entry: { name: trimmedName, formula: '' } };
}

export function createDefinedNamesFromSelection(
  state: State,
  wb: WorkbookHandle,
  source: CreateDefinedNamesSource,
):
  | CreateDefinedNamesResult
  | { ok: false; reason: 'unsupported' | 'empty-selection' | 'engine-failed' } {
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
