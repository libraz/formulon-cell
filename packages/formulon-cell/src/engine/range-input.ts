// Host-supplied ranges and cells given either numerically or as A1 text.
// Declarative options (policy editable cells, viewport selectable lists) keep
// the text and resolve its sheet name on every check, so a qualified range
// follows its sheet through reorders; unqualified text there matches every
// sheet. One-shot actions resolve unqualified text on the active sheet.

import { type ParsedRangeRef, parseA1RangeText, sheetIndexByName } from './range-resolver.js';
import type { Addr, Range } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

/** A range as a numeric `Range` or as A1 text: `B2:B11`, `$B$2`, `B:D`,
 *  `2:11`, `Sheet1!B2:B11`, `'My Sheet'!B2`. */
export type RangeInput = Range | string;

/** A cell as a numeric `Addr` or as A1 text: `B2`, `$B$2`, `Sheet1!B2`. */
export type CellInput = Addr | string;

/** A compiled `RangeInput`: numeric ranges pass through, text keeps its
 *  sheet name (or `null` for every sheet) until it is matched. */
export type RangePattern =
  | { readonly kind: 'range'; readonly range: Range }
  | { readonly kind: 'text'; readonly ref: ParsedRangeRef };

/** Compile a range input, throwing on text that is not an A1 range. */
export function compileRangeInput(input: RangeInput, label: string): RangePattern {
  if (typeof input !== 'string') return { kind: 'range', range: input };
  const ref = parseA1RangeText(input);
  if (!ref) throw new Error(`${label}: "${input}" is not an A1 range`);
  return { kind: 'text', ref };
}

/** Throw when a compiled pattern names a sheet the workbook does not have. */
export function assertPatternSheet(pattern: RangePattern, wb: WorkbookHandle, label: string): void {
  if (pattern.kind !== 'text' || pattern.ref.sheetName === null) return;
  if (sheetIndexByName(wb, pattern.ref.sheetName) < 0) {
    throw new Error(`${label}: sheet "${pattern.ref.sheetName}" does not exist`);
  }
}

/** The pattern's rectangle on `sheet`, or null when it targets another sheet.
 *  A sheet name that no longer exists matches nothing. */
export function patternOnSheet(
  pattern: RangePattern,
  sheet: number,
  wb: WorkbookHandle,
): Range | null {
  if (pattern.kind === 'range') return pattern.range.sheet === sheet ? pattern.range : null;
  const { sheetName, r0, c0, r1, c1 } = pattern.ref;
  if (sheetName !== null && sheetIndexByName(wb, sheetName) !== sheet) return null;
  return { sheet, r0, c0, r1, c1 };
}

export function patternContainsAddr(
  pattern: RangePattern,
  addr: Addr,
  wb: WorkbookHandle,
): boolean {
  const range = patternOnSheet(pattern, addr.sheet, wb);
  return (
    range !== null &&
    addr.row >= range.r0 &&
    addr.row <= range.r1 &&
    addr.col >= range.c0 &&
    addr.col <= range.c1
  );
}

/** Resolve a range input to a numeric range; unqualified text lands on
 *  `activeSheet`. Throws on bad text or an unknown sheet name. */
export function resolveRangeInput(
  input: RangeInput,
  wb: WorkbookHandle,
  activeSheet: number,
  label: string,
): Range {
  const pattern = compileRangeInput(input, label);
  if (pattern.kind === 'range') return pattern.range;
  assertPatternSheet(pattern, wb, label);
  const { sheetName, r0, c0, r1, c1 } = pattern.ref;
  const sheet = sheetName === null ? activeSheet : sheetIndexByName(wb, sheetName);
  return { sheet, r0, c0, r1, c1 };
}

/** Resolve a cell input to an address; unqualified text lands on
 *  `activeSheet`. Throws on text that is not a single A1 cell. */
export function resolveCellInput(
  input: CellInput,
  wb: WorkbookHandle,
  activeSheet: number,
  label: string,
): Addr {
  if (typeof input !== 'string') return input;
  const range = resolveRangeInput(input, wb, activeSheet, label);
  if (range.r0 !== range.r1 || range.c0 !== range.c1) {
    throw new Error(`${label}: "${input}" is not a single cell`);
  }
  return { sheet: range.sheet, row: range.r0, col: range.c0 };
}
