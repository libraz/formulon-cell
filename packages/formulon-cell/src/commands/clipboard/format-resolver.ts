import { addrKey } from '../../engine/address.js';
import { effectiveCellFormatWithPending } from '../../engine/cell-format-sync.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import type { CellFormat, State } from '../../store/store.js';
import type { ClipboardFormatResolver } from './snapshot.js';

/** A fresh resolver for one copy operation. It shares engine XF projections
 *  across cells and address results across snapshot/HTML materialization. */
export function createCopyFormatResolver(
  wb: WorkbookHandle,
  state: State,
): ClipboardFormatResolver {
  const effectiveByXf = new Map<string, Partial<CellFormat> | null>();
  const byAddress = new Map<string, ReturnType<ClipboardFormatResolver>>();

  return (addr) => {
    const key = addrKey(addr);
    if (byAddress.has(key)) return byAddress.get(key);
    const format =
      wb.capabilities.effectiveStyle && !wb.isStub
        ? effectiveCellFormatWithPending(wb, state, addr, effectiveByXf)
        : state.format.formats.get(key);
    byAddress.set(key, format);
    return format;
  };
}
