import {
  addrKey,
  colFromLetters,
  formatA1Range,
  MAX_COL,
  MAX_ROW,
  quoteSheetName,
} from '../engine/address.js';
import { rangeForRef } from '../engine/range-resolver.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import { formatCell, formulaTextLiteral, fromEngineValue } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { isSingleCell } from '../store/selection-geometry.js';
import type { SpreadsheetStore, State } from '../store/store.js';
import { blockDataTableCells } from './data-table-guard.js';
import { adjustFormulaForCutPasteMove } from './formula-refs.js';
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

/** The selection's edges whose labels name the rest of each row or column. */
export interface CreateDefinedNamesEdges {
  top: boolean;
  left: boolean;
  bottom: boolean;
  right: boolean;
}

/** A workbook-scoped name Create from Selection defines, and the range it refers to. */
export interface PlannedDefinedName {
  name: string;
  formula: string;
}

type DefinedNamesSnapshot = Map<string, DefinedNameEntry>;

const MAX_CREATE_DEFINED_NAMES = 10_000;

/**
 * Whether `name` collides with a cell reference and so may not be used as a
 * defined name. Covers A1-style addresses within the sheet grid (up to column
 * XFD / row 1048576) and R1C1-style tokens (R, C, RC, R1C1, …).
 */
const looksLikeCellRef = (name: string): boolean =>
  isA1CellRef(name) || /^R[0-9]*C[0-9]*$/i.test(name);

/** Whether `name` is an A1 address inside the sheet grid (`XFD1048576`, not `XFE1`). */
const isA1CellRef = (name: string): boolean => {
  const a1 = name.toUpperCase().match(/^([A-Z]{1,3})([0-9]{1,7})$/);
  if (!a1) return false;
  const col = colFromLetters(a1[1] ?? '');
  const row = Number.parseInt(a1[2] ?? '', 10) - 1;
  return col >= 0 && col <= MAX_COL && row >= 0 && row <= MAX_ROW;
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

/** A name that starts like a numbered R1C1 reference (`R1C1x`, `R2Cx`), or
 *  in the `_xlnm.` namespace Excel keeps for built-in names, is refused
 *  even though it is not a reference itself. */
const RESERVED_PREFIX = /^(?:R\d+C\d*|R\d*C\d+|_xlnm\.)/i;

const RESERVED_WORD = /^(?:true|false|[rc])$/i;

/** Longest defined name; sanitizing stops one short so a trailing `_` still fits. */
const MAX_DEFINED_NAME_LENGTH = 255;

const isReservedName = (name: string): boolean =>
  RESERVED_WORD.test(name) || RESERVED_PREFIX.test(name) || looksLikeCellRef(name);

/**
 * Validate a defined name against spreadsheet naming rules. Surrounding
 * whitespace is trimmed and fullwidth ASCII is folded before checking, so
 * `Ａ１` is the cell reference A1. The name may not be TRUE/FALSE, a bare
 * `R`/`C`, an A1 or R1C1 reference, or longer than 255 characters; ASCII
 * punctuation, whitespace, currency symbols and `、` are rejected anywhere.
 */
export function isValidDefinedName(raw: string): boolean {
  const name = foldName(raw.trim());
  if (!name || name.length > MAX_DEFINED_NAME_LENGTH) return false;
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
  return findDefinedName(wb, name, localSheetId) !== undefined;
}

/** The text label of a cell; a number, a boolean, an error or a blank cell labels nothing. */
const labelText = (state: State, sheet: number, row: number, col: number): string => {
  const value = state.data.cells.get(addrKey({ sheet, row, col }))?.value;
  return value?.kind === 'text' ? value.value : '';
};

/**
 * Turn label text into a defined name the way Create from Selection does, or `''` when the label
 * gives none. Surrounding whitespace and a leading `=` are dropped, and whitespace runs and other
 * characters a name cannot hold become `_`. A label that reads as an A1 cell, TRUE, FALSE or a
 * bare R or C takes a trailing `_`; one shaped like an R1C1 reference, or starting with a digit,
 * a leading `_`.
 */
export const sanitizeName = (raw: string): string => {
  const folded = foldName(raw.trim()).replace(/^=+/, '').replace(/\s+/g, '_');
  let body = '';
  for (const ch of folded) body += isNameChar(ch, false) ? ch : '_';
  body = body.slice(0, MAX_DEFINED_NAME_LENGTH - 1);
  if (!body) return '';
  if (isA1CellRef(body) || RESERVED_WORD.test(body)) return `${body}_`;
  const first = [...body][0] ?? '';
  return isNameChar(first, true) && !isReservedName(body) ? body : `_${body}`;
};

const snapshotKey = (entry: DefinedNameEntry): string =>
  `${entry.localSheetId}:${entry.name.toLowerCase()}`;

const attrsOf = (entry: DefinedNameEntry): DefinedNameAttrs => ({
  comment: entry.comment ?? '',
  hidden: !!entry.hidden,
});

const captureDefinedNamesSnapshot = (wb: WorkbookHandle): DefinedNamesSnapshot =>
  new Map(listDefinedNames(wb).map((entry) => [snapshotKey(entry), { ...entry }]));

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
  for (const entry of listDefinedNames(wb)) {
    if (!snap.has(snapshotKey(entry))) wb.setDefinedNameEntry(entry.name, '', entry.localSheetId);
  }
  for (const entry of snap.values()) {
    wb.setDefinedNameEntry(entry.name, entry.formula, entry.localSheetId);
    if (wb.capabilities.definedNameAttrs) {
      wb.setDefinedNameAttrs(entry.name, entry.localSheetId, attrsOf(entry));
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

/**
 * Re-point defined names at cells a cut-paste moved: a name whose references
 * lie entirely inside `source` follows the move, one that only partly overlaps
 * it is left alone. Recorded as one undoable step. Returns true when a name
 * changed.
 */
export function moveDefinedNamesForCut(
  history: History | null,
  wb: WorkbookHandle,
  source: Range,
  destination: Range,
): boolean {
  if (!wb.capabilities.definedNameMutate) return false;
  const sheetNames = Array.from({ length: wb.sheetCount }, (_, i) => wb.sheetName(i));
  // Unqualified references in a name have no home sheet (formulaSheet -1), so only
  // qualified ones move, and moved ones always render with an explicit sheet.
  const context = {
    sourceSheet: source.sheet,
    destinationSheet: destination.sheet,
    formulaSheet: -1,
    outputSheet: -1,
    sheetNames,
  };
  let changed = false;
  recordDefinedNamesChange(history, wb, () => {
    for (const entry of listDefinedNames(wb)) {
      // The ref rewriter only tokenizes formula text, which starts with `=`.
      const hasEquals = entry.formula.startsWith('=');
      const moved = adjustFormulaForCutPasteMove(
        hasEquals ? entry.formula : `=${entry.formula}`,
        source,
        destination,
        context,
      );
      const next = hasEquals ? moved : moved.slice(1);
      if (next === entry.formula) continue;
      if (wb.setDefinedNameEntry(entry.name, next, entry.localSheetId)) changed = true;
      if (wb.capabilities.definedNameAttrs && (entry.comment || entry.hidden)) {
        wb.setDefinedNameAttrs(entry.name, entry.localSheetId, attrsOf(entry));
      }
    }
  });
  if (changed) wb.recalcAuto();
  return changed;
}

/** A stored defined-name formula as name lists show it: with a leading `=`. */
export function definedNameDisplayFormula(formula: string): string {
  return formula && !formula.startsWith('=') ? `=${formula}` : formula;
}

const MAX_VALUE_TEXT = 64;
const NUMBER_LITERAL = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i;
const STRING_LITERAL = /^"(?:[^"]|"")*"$/;

const arrayElementText = (v: CellValue): string =>
  v.kind === 'text' ? formulaTextLiteral(v.value) : v.kind === 'blank' ? '""' : formatCell(v);

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
    if (isSingleCell(range)) {
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
  if (!wb.capabilities.formulaTextEvaluation) return '';
  try {
    const result = wb.evaluateFormulaText({ sheet: ownSheet, row: 0, col: 0 }, `=${body}`);
    if (!result) return '';
    const value = fromEngineValue(result);
    return truncateValueText(
      value.kind === 'text' ? formulaTextLiteral(value.value) : formatCell(value),
    );
  } catch {
    return '';
  }
}

/** Snapshot workbook- and sheet-scoped defined names into a stable array. */
export function listDefinedNames(wb: WorkbookHandle): DefinedNameEntry[] {
  return [...wb.definedNames()];
}

/** The defined name equal to `name` (ignoring case and width) in the given scope. */
function findDefinedName(
  wb: WorkbookHandle,
  name: string,
  localSheetId: number,
): DefinedNameEntry | undefined {
  return listDefinedNames(wb).find(
    (e) => e.localSheetId === localSheetId && isSameDefinedName(e.name, name),
  );
}

/** Whether the workbook has a name the Name Manager lists: any scope, hidden names excluded. */
export function hasListedDefinedNames(wb: WorkbookHandle): boolean {
  for (const entry of wb.definedNames()) if (!entry.hidden) return true;
  return false;
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
  // Redefining a name under another casing renames it to the new spelling
  // (Excel: adding `DUP` over `dup` stores `DUP`); its note and hidden flag stay.
  const previous = listDefinedNames(wb).find(
    (e) =>
      e.localSheetId === localSheetId &&
      isSameDefinedName(e.name, trimmedName) &&
      e.name !== trimmedName,
  );
  if (previous && !wb.setDefinedNameEntry(previous.name, '', localSheetId)) {
    return { ok: false, reason: 'engine-failed' };
  }
  if (!wb.setDefinedNameEntry(trimmedName, trimmedFormula, localSheetId)) {
    if (previous) wb.setDefinedNameEntry(previous.name, previous.formula, localSheetId);
    return { ok: false, reason: 'engine-failed' };
  }
  const keptAttrs: DefinedNameAttrs = {
    ...(previous?.comment !== undefined ? { comment: previous.comment } : {}),
    ...(previous?.hidden !== undefined ? { hidden: previous.hidden } : {}),
    ...attrs,
  };
  if (
    wb.capabilities.definedNameAttrs &&
    (keptAttrs.comment !== undefined || keptAttrs.hidden !== undefined) &&
    !wb.setDefinedNameAttrs(trimmedName, localSheetId, keptAttrs)
  ) {
    return { ok: false, reason: 'engine-failed' };
  }
  // Defining/redefining a name changes what `=MyRange` resolves to — recompute
  // so dependent cells and a subsequent save reflect the new target.
  wb.recalcAuto();
  // The entry is the name as the engine stored it, read back: its formula may be respelled.
  const written = findDefinedName(wb, trimmedName, localSheetId);
  return {
    ok: true,
    entry: {
      name: written?.name ?? trimmedName,
      formula: written?.formula ?? trimmedFormula.replace(/^=/, ''),
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

/**
 * The names Create from Selection defines over `range` from the labels on each checked edge. An
 * edge's names leave out the other checked edges, so with Top row and Left column checked the
 * column names start after the left column and the row names after the top row. A label that
 * gives no name is skipped; where two labels give the same name (ignoring case and width), the
 * later one wins with its own spelling, the edges taken top, left, bottom, right.
 */
export function planDefinedNamesFromSelection(
  state: State,
  wb: WorkbookHandle,
  range: Range,
  edges: CreateDefinedNamesEdges,
): { ok: true; names: PlannedDefinedName[] } | { ok: false; reason: 'too-large' } {
  const { sheet } = range;
  const sheetName = quoteSheetName(wb.sheetName(sheet));
  const r0 = edges.top ? range.r0 + 1 : range.r0;
  const r1 = edges.bottom ? range.r1 - 1 : range.r1;
  const c0 = edges.left ? range.c0 + 1 : range.c0;
  const c1 = edges.right ? range.c1 - 1 : range.c1;
  const planned = new Map<string, PlannedDefinedName>();
  const add = (label: string, refers: Omit<Range, 'sheet'>): void => {
    const name = sanitizeName(label);
    if (!name) return;
    const key = nameKey(name);
    planned.delete(key);
    planned.set(key, {
      name,
      formula: `=${sheetName}!${formatA1Range(refers, { absolute: true })}`,
    });
  };
  // `labels` is the label row or column; the names refer from `from` to `to` across it.
  const sides = [
    { checked: edges.top, across: 'cols', labels: range.r0, from: range.r0 + 1, to: range.r1 },
    { checked: edges.left, across: 'rows', labels: range.c0, from: range.c0 + 1, to: range.c1 },
    { checked: edges.bottom, across: 'cols', labels: range.r1, from: range.r0, to: range.r1 - 1 },
    { checked: edges.right, across: 'rows', labels: range.c1, from: range.c0, to: range.c1 - 1 },
  ] as const;
  for (const side of sides) {
    if (!side.checked || side.from > side.to) continue;
    const count = side.across === 'cols' ? c1 - c0 + 1 : r1 - r0 + 1;
    if (count > MAX_CREATE_DEFINED_NAMES) return { ok: false, reason: 'too-large' };
    if (side.across === 'cols') {
      for (let col = c0; col <= c1; col += 1) {
        const label = labelText(state, sheet, side.labels, col);
        add(label, { r0: side.from, c0: col, r1: side.to, c1: col });
      }
    } else {
      for (let row = r0; row <= r1; row += 1) {
        const label = labelText(state, sheet, row, side.labels);
        add(label, { r0: row, c0: side.from, r1: row, c1: side.to });
      }
    }
  }
  return { ok: true, names: [...planned.values()] };
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
    if (!isSameDefinedName(definedName.name, trimmed)) continue;
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
  if (blockDataTableCells(wb, [addr])) return null;
  const formula = `=${entry.name}`;
  wb.setFormula(addr, formula);
  return { addr, formula };
}
