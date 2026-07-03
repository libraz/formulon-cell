import { addrKey } from '../../engine/address.js';
import type { Addr, Range } from '../../engine/types.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import { type CellFormat, mutators, type SpreadsheetStore, type State } from '../../store/store.js';
import { adjustFormulaForCutPasteMove } from '../formula-refs.js';
import { isCellWritable } from '../protection.js';
import { shiftFormulaRefs } from '../refs.js';
import type { ClipboardCell, ClipboardSnapshot } from './snapshot.js';

export type PasteWhat =
  | 'all'
  | 'values'
  | 'formulas'
  | 'formats'
  | 'formulas-and-numfmt'
  | 'values-and-numfmt';

export type PasteOperation = 'none' | 'add' | 'subtract' | 'multiply' | 'divide';

export interface PasteSpecialOptions {
  what: PasteWhat;
  operation: PasteOperation;
  skipBlanks: boolean;
  transpose: boolean;
}

export interface PasteSpecialResult {
  writtenRange: Range;
  /** Arithmetic operations that produced NaN/Infinity and were written as
   *  static spreadsheet error values. */
  skippedNonFiniteOperations: number;
}

const numericValue = (cell: ClipboardCell | undefined): number | null => {
  if (!cell) return null;
  return cell.value.kind === 'number' ? cell.value.value : null;
};

const writeClipboardValue = (wb: WorkbookHandle, addr: Addr, src: ClipboardCell): void => {
  switch (src.value.kind) {
    case 'number':
      wb.setNumber(addr, src.value.value);
      break;
    case 'text':
      wb.setText(addr, src.value.value);
      break;
    case 'bool':
      wb.setBool(addr, src.value.value);
      break;
    case 'blank':
      wb.setBlank(addr);
      break;
    case 'error':
      wb.setError(addr, src.value.code);
      break;
  }
};

const existingNumeric = (state: State, sheet: number, row: number, col: number): number => {
  const cell = state.data.cells.get(addrKey({ sheet, row, col }));
  if (!cell) return 0;
  if (cell.value.kind === 'number') return cell.value.value;
  return 0;
};

const combine = (op: PasteOperation, dest: number, src: number): number => {
  switch (op) {
    case 'add':
      return dest + src;
    case 'subtract':
      return dest - src;
    case 'multiply':
      return dest * src;
    case 'divide':
      return src === 0 ? Number.NaN : dest / src;
    default:
      return src;
  }
};

const errorCodeForNonFiniteOperation = (op: PasteOperation): number => (op === 'divide' ? 1 : 5); // #DIV/0! for divide, #NUM! for arithmetic overflow.

const wantsValues = (what: PasteWhat): boolean =>
  what === 'all' || what === 'values' || what === 'values-and-numfmt';
const wantsFormulas = (what: PasteWhat): boolean =>
  what === 'all' || what === 'formulas' || what === 'formulas-and-numfmt';
const wantsFormats = (what: PasteWhat): boolean => what === 'all' || what === 'formats';
const wantsNumFmt = (what: PasteWhat): boolean =>
  what === 'all' ||
  what === 'formats' ||
  what === 'values-and-numfmt' ||
  what === 'formulas-and-numfmt';

/**
 * Apply a clipboard snapshot to the destination starting at `state.selection.active`,
 * filtered by the spreadsheet-style "Paste Special" options. Returns the range that was
 * actually written. Caller is responsible for refreshing the cached cell map.
 */
export function pasteSpecial(
  state: State,
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  snap: ClipboardSnapshot,
  opt: PasteSpecialOptions,
): PasteSpecialResult | null {
  const origin: Addr = state.selection.active;
  const sheet = origin.sheet;

  const destRows = opt.transpose ? snap.cols : snap.rows;
  const destCols = opt.transpose ? snap.rows : snap.cols;

  // Pre-compute format patches we'll merge into the format slice in one pass.
  const formatWrites: { key: string; format: CellFormat | null }[] = [];
  let skippedNonFiniteOperations = 0;

  for (let dr = 0; dr < destRows; dr += 1) {
    for (let dc = 0; dc < destCols; dc += 1) {
      const sr = opt.transpose ? dc : dr;
      const sc = opt.transpose ? dr : dc;
      const src = snap.cells[sr]?.[sc];
      if (!src) continue;
      const isBlankSrc = src.value.kind === 'blank' && !src.formula && !src.format;
      if (opt.skipBlanks && isBlankSrc) continue;

      const row = origin.row + dr;
      const col = origin.col + dc;
      const addr: Addr = { sheet, row, col };
      // Sheet protection — silently skip locked destinations (spreadsheet parity).
      if (!isCellWritable(state, addr)) continue;

      // Layer 1: values / formulas.
      // A Paste Special arithmetic operation always combines by VALUE. When one
      // is active it takes precedence over formula-pasting, using the source's
      // computed number even if the source cell is a formula — otherwise an
      // "Add" over a formula source silently pastes the formula and drops the
      // operation. Formats-only pastes carry no value, so they never
      // operate.
      const operating =
        opt.operation !== 'none' && (wantsValues(opt.what) || wantsFormulas(opt.what));
      const shouldPasteFormula = Boolean(src.formula && wantsFormulas(opt.what) && !operating);
      if (operating) {
        const srcNum = numericValue(src);
        if (srcNum !== null) {
          const dest = existingNumeric(state, sheet, row, col);
          const result = combine(opt.operation, dest, srcNum);
          if (Number.isFinite(result)) {
            wb.setNumber(addr, result);
          } else {
            wb.setError(addr, errorCodeForNonFiniteOperation(opt.operation));
            skippedNonFiniteOperations += 1;
          }
        }
        // Non-numeric source cells leave the destination unchanged (parity).
      } else if (shouldPasteFormula && src.formula) {
        // Cut moves cells: paste formulas verbatim (Excel keeps references
        // intact on a move). Copy re-anchors relative refs by the paste offset.
        if (snap.mode === 'cut') {
          wb.setFormula(addr, src.formula);
        } else {
          const sourceRow = snap.range.r0 + sr;
          const sourceCol = snap.range.c0 + sc;
          wb.setFormula(addr, shiftFormulaRefs(src.formula, row - sourceRow, col - sourceCol));
        }
      } else if (wantsValues(opt.what)) {
        if (src.value.kind !== 'blank' || !isBlankSrc) writeClipboardValue(wb, addr, src);
      }

      // Layer 2: formats
      const fmt = src.format;
      if (wantsFormats(opt.what)) {
        // A full "Formats" paste copies the source's *absence* of formatting
        // too: an unformatted source cell clears the destination format rather
        // than leaving a stale one behind (spreadsheet parity).
        formatWrites.push({
          key: addrKey(addr),
          format: fmt ? { ...fmt, borders: fmt.borders ? { ...fmt.borders } : undefined } : null,
        });
      } else if (wantsNumFmt(opt.what) && fmt?.numFmt) {
        // Number format only — cherry-pick.
        formatWrites.push({ key: addrKey(addr), format: { numFmt: fmt.numFmt } });
      }
    }
  }

  if (formatWrites.length > 0) {
    store.setState((s) => {
      const formats = new Map(s.format.formats);
      for (const { key, format } of formatWrites) {
        if (format === null) {
          formats.delete(key);
        } else if (wantsFormats(opt.what)) {
          formats.set(key, format);
        } else {
          // For cherry-picks (number format only), merge into the existing cell format.
          const existing = formats.get(key) ?? {};
          formats.set(key, { ...existing, ...format });
        }
      }
      return { ...s, format: { ...s.format, formats } };
    });
  }
  if (skippedNonFiniteOperations > 0) {
    console.warn(
      `formulon-cell: paste special wrote ${skippedNonFiniteOperations} non-finite arithmetic result(s) as static error value(s)`,
    );
  }

  // Move active selection to the written range.
  const writtenRange: Range = {
    sheet,
    r0: origin.row,
    c0: origin.col,
    r1: origin.row + destRows - 1,
    c1: origin.col + destCols - 1,
  };
  mutators.setActive(store, { sheet, row: writtenRange.r0, col: writtenRange.c0 });
  if (writtenRange.r0 !== writtenRange.r1 || writtenRange.c0 !== writtenRange.c1) {
    mutators.extendRangeTo(store, { sheet, row: writtenRange.r1, col: writtenRange.c1 });
  }
  if (snap.mode === 'cut') {
    updateExternalRefsForCutPaste(wb, snap.range, writtenRange);
  }
  return { writtenRange, skippedNonFiniteOperations };
}

function updateExternalRefsForCutPaste(
  wb: WorkbookHandle,
  source: Range,
  writtenRange: Range,
): void {
  if (source.sheet !== writtenRange.sheet) return;
  for (const entry of Array.from(wb.cells(writtenRange.sheet))) {
    if (!entry.formula) continue;
    if (
      entry.addr.row >= writtenRange.r0 &&
      entry.addr.row <= writtenRange.r1 &&
      entry.addr.col >= writtenRange.c0 &&
      entry.addr.col <= writtenRange.c1
    ) {
      continue;
    }
    const next = adjustFormulaForCutPasteMove(entry.formula, source, writtenRange);
    if (next !== entry.formula) wb.setFormula(entry.addr, next);
  }
}
