import type { ConditionalCellOverlay } from '../render/conditional.js';
import { addrKey } from './address.js';
import type { WorkbookHandle } from './workbook-handle.js';

/** CF match kind ordinals — mirror of `formulon::cf::CFMatchKind`. */
const KIND_COLOR_SCALE = 1;
const KIND_DATA_BAR = 2;

const rgba = (c: { r: number; g: number; b: number; a: number }): string =>
  c.a >= 255
    ? `rgb(${c.r}, ${c.g}, ${c.b})`
    : `rgba(${c.r}, ${c.g}, ${c.b}, ${(c.a / 255).toFixed(3)})`;

/**
 * Evaluate engine-side CF rules over `[(firstRow, firstCol), (lastRow, lastCol)]`
 * on `sheet` and lift the result into `ConditionalCellOverlay` shape so it can
 * be merged with the JS-side overlay map.
 *
 * Today only ColorScale (fill) and DataBar (bar + barColor) lift cleanly —
 * DifferentialFormat needs the dxf table (not exposed yet) and IconSet needs
 * a glyph renderer. Those matches are silently dropped.
 *
 * Returns an empty map when the engine doesn't expose `evaluateCfRange`.
 */
export function evaluateCfFromEngine(
  wb: WorkbookHandle,
  sheet: number,
  firstRow: number,
  firstCol: number,
  lastRow: number,
  lastCol: number,
  todaySerial: number = Number.NaN,
): Map<string, ConditionalCellOverlay> {
  const out = new Map<string, ConditionalCellOverlay>();
  if (!wb.capabilities.conditionalFormat) return out;
  const cells = wb.evaluateCfRange(sheet, firstRow, firstCol, lastRow, lastCol, todaySerial);
  for (const cell of cells) {
    const key = addrKey({ sheet, row: cell.row, col: cell.col });
    const overlay: ConditionalCellOverlay = out.get(key) ?? {};
    // Iterate matches in priority order — engine returns them sorted by
    // priority, so later writes win for fields like `fill` (regular CF
    // semantics: highest priority match overrides).
    for (const m of cell.matches) {
      if (m.kind === KIND_COLOR_SCALE) {
        overlay.fill = rgba(m.color);
      } else if (m.kind === KIND_DATA_BAR) {
        overlay.bar = Math.max(0, Math.min(1, m.barLengthPct / 100));
        overlay.barColor = rgba(m.barFill);
      }
    }
    if (Object.keys(overlay).length > 0) out.set(key, overlay);
  }
  return out;
}
