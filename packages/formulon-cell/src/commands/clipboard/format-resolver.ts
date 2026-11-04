import { addrKey } from '../../engine/address.js';
import { effectiveCellFormat } from '../../engine/cell-format-projection.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import { styleElementFormatSource } from '../../render/grid/table-format.js';
import type { CellFormat, State } from '../../store/store.js';
import type { ClipboardFormatResolver } from './snapshot.js';

/** `format` over what a table or PivotTable style paints at the cell: a copy
 *  carries the style as direct formatting, as the desktop pastes it. */
function withStyle(
  format: Partial<CellFormat> | null | undefined,
  style: CellFormat | undefined,
): Partial<CellFormat> | null | undefined {
  if (!style) return format;
  const out: Partial<CellFormat> = { ...style };
  for (const [key, value] of Object.entries(format ?? {})) {
    if (value !== undefined) (out as Record<string, unknown>)[key] = value;
  }
  if (style.borders && format?.borders) {
    const borders = { ...style.borders };
    for (const [side, edge] of Object.entries(format.borders)) {
      if (edge !== undefined) (borders as Record<string, unknown>)[side] = edge;
    }
    out.borders = borders;
  }
  return out;
}

/** A fresh resolver for one copy operation. It shares engine XF projections
 *  across cells and address results across snapshot/HTML materialization. */
export function createCopyFormatResolver(
  wb: WorkbookHandle,
  state: State,
): ClipboardFormatResolver {
  const effectiveByXf = new Map<string, Partial<CellFormat> | null>();
  const byAddress = new Map<string, ReturnType<ClipboardFormatResolver>>();
  const styleAt = styleElementFormatSource(state, wb);

  return (addr) => {
    const key = addrKey(addr);
    if (byAddress.has(key)) return byAddress.get(key);
    const format =
      wb.capabilities.effectiveStyle && !wb.isStub
        ? effectiveCellFormat(wb, state, addr, effectiveByXf)
        : state.format.formats.get(key);
    const flattened =
      addr.sheet === state.data.sheetIndex
        ? withStyle(format, styleAt(addr.row, addr.col))
        : format;
    byAddress.set(key, flattened);
    return flattened;
  };
}
