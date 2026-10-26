import { colFromLetters, colLetter, formatA1Range, MAX_COL, MAX_ROW } from '../engine/address.js';
import { rangeForRef } from '../engine/range-resolver.js';
import type { Addr, CellValue } from '../engine/types.js';
import { formatCell, fromEngineValue } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore, State } from '../store/store.js';
import type { History } from './history.js';
import { isCellWritable, warnProtected } from './protection.js';

export interface DefinedNameEntry {
  name: string;
  formula: string;
  /** -1 means workbook scope; otherwise a 0-based sheet index. */
  localSheetId: number;
  /** The note the Name Manager shows; absent when empty. */
  comment?: string;
  /** Hidden from the Name Manager and name lists, but still resolvable in formulas. */
  hidden?: boolean;
}

/** Attributes written next to a name's formula; an omitted key is left as is. */
export interface DefinedNameAttrs {
  comment?: string;
  hidden?: boolean;
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

type DefinedNamesSnapshot = Map<string, DefinedNameEntry>;

const MAX_CREATE_DEFINED_NAMES = 10_000;

/**
 * Whether `name` collides with a cell reference and so may not be used as a
 * defined name. Covers A1-style addresses within the sheet grid (up to column
 * XFD / row 1048576) and R1C1-style tokens (R, C, RC, R1C1, …).
 */
const looksLikeCellRef = (name: string): boolean => {
  const up = name.toUpperCase();
  const a1 = up.match(/^([A-Z]{1,3})([0-9]{1,7})$/);
  if (a1) {
    const col = colFromLetters(a1[1] ?? '');
    const row = Number.parseInt(a1[2] ?? '', 10) - 1;
    if (col >= 0 && col <= MAX_COL && row >= 0 && row <= MAX_ROW) return true;
  }
  if (/^R[0-9]*C[0-9]*$/.test(up)) return true;
  return false;
};

/** Fold fullwidth ASCII (U+FF01-FF5E), the wave dash and the fullwidth yen to their ASCII-range forms. */
const foldName = (raw: string): string => {
  let out = '';
  for (const ch of raw) {
    const code = ch.codePointAt(0) ?? 0;
    if (code >= 0xff01 && code <= 0xff5e) out += String.fromCharCode(code - 0xfee0);
    else if (code === 0x301c) out += '~';
    else if (code === 0xffe5) out += '¥';
    else out += ch;
  }
  return out;
};

const FIRST_ASCII = /^[A-Za-z_\\]$/;
const REST_ASCII = /^[A-Za-z0-9_.?\\]$/;
const BLOCKED_NON_ASCII = /^[\s\p{Z}\p{Cc}\p{Sc}、]$/u;

/** Whether `ch` (one code point of folded text) may appear at `first` / later positions. */
const isNameChar = (ch: string, first: boolean): boolean => {
  if ((ch.codePointAt(0) ?? 0) < 0x80) return (first ? FIRST_ASCII : REST_ASCII).test(ch);
  if (BLOCKED_NON_ASCII.test(ch)) return false;
  return !(first && ch === '。');
};

const isReservedName = (name: string): boolean =>
  /^(?:true|false|[rc])$/i.test(name) || looksLikeCellRef(name);

/**
 * Validate a defined name against spreadsheet naming rules. Surrounding
 * whitespace is trimmed and fullwidth ASCII is folded before checking, so
 * `Ａ１` is the cell reference A1. The name may not be TRUE/FALSE, a bare
 * `R`/`C`, an A1 or R1C1 reference, or longer than 255 characters; ASCII
 * punctuation, whitespace, currency symbols and `、` are rejected anywhere.
 */
export function isValidDefinedName(raw: string): boolean {
  const name = foldName(raw.trim());
  if (!name || name.length > 255) return false;
  if (isReservedName(name)) return false;
  let first = true;
  for (const ch of name) {
    if (!isNameChar(ch, first)) return false;
    first = false;
  }
  return true;
}

/** Case- and width-insensitive key of a defined name, for duplicate checks. */
const nameKey = (name: string): string => foldName(name.trim()).toLowerCase();

/** Whether two names are the same defined name (ignoring case and width). */
export const isSameDefinedName = (a: string, b: string): boolean => nameKey(a) === nameKey(b);

/** Whether `wb` already has a name equal to `name` (ignoring case and width) in the given scope. */
export function definedNameExists(wb: WorkbookHandle, name: string, localSheetId = -1): boolean {
  const key = nameKey(name);
  for (const entry of wb.definedNames()) {
    if (entry.localSheetId === localSheetId && nameKey(entry.name) === key) return true;
  }
  return false;
}

const absRangeRef = (r0: number, c0: number, r1: number, c1: number): string =>
  formatA1Range({ r0, c0, r1, c1 }, { absolute: true, collapse: false });

const cellText = (state: State, sheet: number, row: number, col: number): string => {
  const cell = state.data.cells.get(`${sheet}:${row}:${col}`);
  if (!cell) return '';
  if (cell.value.kind === 'text') return cell.value.value;
  if (cell.value.kind === 'number') return String(cell.value.value);
  return '';
};

/** Turn arbitrary label text into a name that passes `isValidDefinedName`. */
export const sanitizeName = (raw: string, fallback: string): string => {
  const folded = foldName(raw.trim()).replace(/\s+/g, '_');
  let body = '';
  for (const ch of folded) body += isNameChar(ch, false) ? ch : '_';
  body = (body || foldName(fallback)).slice(0, 254);
  const prefixed = isNameChar([...body][0] ?? '', true) ? body : `_${body}`;
  // A label like "A1" or "TRUE" is a valid identifier but not a usable name.
  return isReservedName(prefixed) ? `_${prefixed}` : prefixed;
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
  new Map(
    [...wb.definedNames()].map((entry) => [
      `${entry.localSheetId}:${entry.name.toLowerCase()}`,
      { ...entry },
    ]),
  );

const sameDefinedNamesSnapshot = (a: DefinedNamesSnapshot, b: DefinedNamesSnapshot): boolean => {
  if (a.size !== b.size) return false;
  for (const [key, entry] of a) {
    const other = b.get(key);
    if (
      !other ||
      other.name !== entry.name ||
      other.formula !== entry.formula ||
      other.localSheetId !== entry.localSheetId ||
      (other.comment ?? '') !== (entry.comment ?? '') ||
      !!other.hidden !== !!entry.hidden
    ) {
      return false;
    }
  }
  return true;
};

const applyDefinedNamesSnapshot = (wb: WorkbookHandle, snap: DefinedNamesSnapshot): void => {
  for (const entry of [...wb.definedNames()]) {
    const key = `${entry.localSheetId}:${entry.name.toLowerCase()}`;
    if (!snap.has(key)) wb.setDefinedNameEntry(entry.name, '', entry.localSheetId);
  }
  for (const entry of snap.values()) {
    wb.setDefinedNameEntry(entry.name, entry.formula, entry.localSheetId);
    if (wb.capabilities.definedNameAttrs) {
      wb.setDefinedNameAttrs(entry.name, entry.localSheetId, {
        comment: entry.comment ?? '',
        hidden: !!entry.hidden,
      });
    }
  }
  // Undo/redo of a name change must recompute dependents too.
  wb.recalcAuto();
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

/** A stored defined-name formula as name lists show it: with a leading `=`. */
export function definedNameDisplayFormula(formula: string): string {
  return formula && !formula.startsWith('=') ? `=${formula}` : formula;
}

const MAX_VALUE_TEXT = 64;
const NUMBER_LITERAL = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i;
const STRING_LITERAL = /^"(?:[^"]|"")*"$/;

const quoteText = (text: string): string => `"${text.replace(/"/g, '""')}"`;

const arrayElementText = (v: CellValue): string =>
  v.kind === 'text' ? quoteText(v.value) : v.kind === 'blank' ? '""' : formatCell(v);

const truncateValueText = (text: string): string =>
  text.length > MAX_VALUE_TEXT ? `${text.slice(0, MAX_VALUE_TEXT)}…` : text;

/**
 * The Name Manager "Value" text of a defined name: constants as written,
 * a single-cell reference as the cell's display value, a multi-cell range in
 * array form (`{"a","b";"c","d"}`), and other formulas as evaluated by the
 * engine. Empty when nothing can be shown.
 */
export function definedNameValueText(wb: WorkbookHandle, entry: DefinedNameEntry): string {
  const body = entry.formula.trim().replace(/^=/, '').trim();
  if (!body) return '';
  if (NUMBER_LITERAL.test(body)) return formatCell({ kind: 'number', value: Number(body) });
  if (STRING_LITERAL.test(body)) return body;
  if (/^(?:true|false)$/i.test(body)) return body.toUpperCase();
  const ownSheet = Math.max(entry.localSheetId, 0);
  const range = rangeForRef(wb, body, ownSheet);
  if (range) {
    const { sheet } = range;
    if (range.r0 === range.r1 && range.c0 === range.c1) {
      return formatCell(wb.getValue({ sheet, row: range.r0, col: range.c0 }));
    }
    const rows: string[] = [];
    let used = 0;
    for (let row = range.r0; row <= range.r1 && used <= MAX_VALUE_TEXT; row += 1) {
      const cols: string[] = [];
      for (let col = range.c0; col <= range.c1 && used <= MAX_VALUE_TEXT; col += 1) {
        const text = arrayElementText(wb.getValue({ sheet, row, col }));
        cols.push(text);
        used += text.length + 1;
      }
      rows.push(cols.join(','));
    }
    return truncateValueText(`{${rows.join(';')}}`);
  }
  if (!wb.capabilities.formulaTextEvaluation || typeof wb.evaluateFormulaText !== 'function') {
    return '';
  }
  try {
    const result = wb.evaluateFormulaText({ sheet: ownSheet, row: 0, col: 0 }, `=${body}`);
    if (!result.status.ok) return '';
    const value = fromEngineValue(result.value);
    return truncateValueText(value.kind === 'text' ? quoteText(value.value) : formatCell(value));
  } catch {
    return '';
  }
}

/** Snapshot workbook- and sheet-scoped defined names into a stable array. */
export function listDefinedNames(wb: WorkbookHandle): DefinedNameEntry[] {
  return [...wb.definedNames()];
}

/** Add or replace a workbook- or sheet-scoped defined name. */
export function upsertDefinedName(
  wb: WorkbookHandle,
  name: string,
  formula: string,
  localSheetId = -1,
  attrs: DefinedNameAttrs = {},
): DefinedNameMutationResult {
  const trimmedName = name.trim();
  const trimmedFormula = formula.trim();
  if (!trimmedName) return { ok: false, reason: 'empty-name' };
  if (!isValidDefinedName(trimmedName)) return { ok: false, reason: 'invalid-name' };
  if (!trimmedFormula) return { ok: false, reason: 'empty-formula' };
  if (!wb.capabilities.definedNameMutate) return { ok: false, reason: 'unsupported' };
  if (!wb.setDefinedNameEntry(trimmedName, trimmedFormula, localSheetId)) {
    return { ok: false, reason: 'engine-failed' };
  }
  if (
    wb.capabilities.definedNameAttrs &&
    (attrs.comment !== undefined || attrs.hidden !== undefined) &&
    !wb.setDefinedNameAttrs(trimmedName, localSheetId, attrs)
  ) {
    return { ok: false, reason: 'engine-failed' };
  }
  // Defining/redefining a name changes what `=MyRange` resolves to — recompute
  // so dependent cells and a subsequent save reflect the new target.
  wb.recalcAuto();
  // The engine stores the formula without its leading `=`, as `definedNames()` reads it back.
  const stored = trimmedFormula.replace(/^=/, '');
  const written = [...wb.definedNames()].find(
    (e) => e.localSheetId === localSheetId && isSameDefinedName(e.name, trimmedName),
  );
  return {
    ok: true,
    entry: {
      name: trimmedName,
      formula: stored,
      localSheetId,
      ...(written?.comment ? { comment: written.comment } : {}),
      ...(written?.hidden ? { hidden: true } : {}),
    },
  };
}

/** Remove a workbook- or sheet-scoped defined name. */
export function deleteDefinedName(
  wb: WorkbookHandle,
  name: string,
  localSheetId = -1,
): DefinedNameDeleteResult {
  const trimmedName = name.trim();
  if (!trimmedName) return { ok: false, reason: 'empty-name' };
  if (!wb.capabilities.definedNameMutate) return { ok: false, reason: 'unsupported' };
  if (!wb.setDefinedNameEntry(trimmedName, '', localSheetId)) {
    return { ok: false, reason: 'engine-failed' };
  }
  // Removing a name usually turns `=MyRange` into #NAME? — recompute so the
  // change is reflected in dependents and on save.
  wb.recalcAuto();
  return { ok: true, entry: { name: trimmedName, formula: '', localSheetId } };
}

export function createDefinedNamesFromSelection(
  state: State,
  wb: WorkbookHandle,
  source: CreateDefinedNamesSource,
):
  | CreateDefinedNamesResult
  | { ok: false; reason: 'unsupported' | 'empty-selection' | 'too-large' | 'engine-failed' } {
  if (!wb.capabilities.definedNameMutate) return { ok: false, reason: 'unsupported' };
  const sheet = state.selection.range.sheet;
  const r = state.selection.range;
  const used = new Set([...wb.definedNames()].map((entry) => entry.name.toLowerCase()));
  const entries: DefinedNameEntry[] = [];

  if (source === 'top-row') {
    if (r.r0 >= r.r1) return { ok: false, reason: 'empty-selection' };
    if (r.c1 - r.c0 + 1 > MAX_CREATE_DEFINED_NAMES) {
      return { ok: false, reason: 'too-large' };
    }
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
    if (r.c1 - r.c0 + 1 > MAX_CREATE_DEFINED_NAMES) {
      return { ok: false, reason: 'too-large' };
    }
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
    if (r.r1 - r.r0 + 1 > MAX_CREATE_DEFINED_NAMES) {
      return { ok: false, reason: 'too-large' };
    }
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
  if (r.r1 - r.r0 + 1 > MAX_CREATE_DEFINED_NAMES) {
    return { ok: false, reason: 'too-large' };
  }
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
  const activeSheet = state.selection.active.sheet;
  let entry: DefinedNameEntry | null = null;
  let workbookScoped: DefinedNameEntry | null = null;
  for (const definedName of wb.definedNames()) {
    if (definedName.name.toLowerCase() !== trimmed.toLowerCase()) continue;
    if (definedName.localSheetId === activeSheet) {
      entry = definedName;
      break;
    }
    if (definedName.localSheetId === -1 && workbookScoped === null) workbookScoped = definedName;
  }
  entry ??= workbookScoped;
  if (!entry) return null;
  const addr = state.selection.active;
  if (store && !isCellWritable(store.getState(), addr)) {
    warnProtected(addr);
    return null;
  }
  const formula = `=${entry.name}`;
  wb.setFormula(addr, formula);
  return { addr, formula };
}
