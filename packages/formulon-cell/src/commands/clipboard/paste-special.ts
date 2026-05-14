import { addrKey } from '../../engine/address.js';
import type { Addr, Range } from '../../engine/types.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import { type CellFormat, mutators, type SpreadsheetStore, type State } from '../../store/store.js';
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
      // No public way to write an error; skip — pasting errors is rare.
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

      // Layer 1: values / formulas
      const shouldPasteFormula = Boolean(src.formula && wantsFormulas(opt.what));
      if (wantsValues(opt.what) && !shouldPasteFormula) {
        const srcNum = numericValue(src);
        if (opt.operation !== 'none' && srcNum !== null) {
          const dest = existingNumeric(state, sheet, row, col);
          const result = combine(opt.operation, dest, srcNum);
          if (Number.isFinite(result)) wb.setNumber(addr, result);
        } else if (opt.operation !== 'none') {
          // Desktop spreadsheets Paste Special arithmetic operations ignore non-numeric
          // source cells instead of overwriting the destination with text.
        } else {
          if (src.value.kind !== 'blank' || !isBlankSrc) writeClipboardValue(wb, addr, src);
        }
      } else if (shouldPasteFormula && src.formula) {
        const sourceRow = snap.range.r0 + sr;
        const sourceCol = snap.range.c0 + sc;
        wb.setFormula(addr, shiftFormulaRefs(src.formula, row - sourceRow, col - sourceCol));
      }

      // Layer 2: formats
      const fmt = src.format;
      if (fmt) {
        if (wantsFormats(opt.what)) {
          formatWrites.push({
            key: addrKey(addr),
            format: { ...fmt, borders: fmt.borders ? { ...fmt.borders } : undefined },
          });
        } else if (wantsNumFmt(opt.what) && fmt.numFmt) {
          // Number format only — cherry-pick.
          formatWrites.push({ key: addrKey(addr), format: { numFmt: fmt.numFmt } });
        }
      }
    }
  }

  if (formatWrites.length > 0) {
    store.setState((s) => {
      const formats = new Map(s.format.formats);
      for (const { key, format } of formatWrites) {
        if (format === null) {
          formats.delete(key);
        } else {
          // For 'formats' wholesale-replace; for cherry-picks merge.
          const existing = formats.get(key) ?? {};
          formats.set(key, { ...existing, ...format });
        }
      }
      return { ...s, format: { formats } };
    });
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
  return { writtenRange };
}
