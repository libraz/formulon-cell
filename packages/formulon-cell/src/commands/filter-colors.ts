import { addrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import { evaluateConditional } from '../render/conditional.js';
import type { ConditionalIconRef } from '../store/types/conditional-rule.js';
import type { State } from '../store/types.js';

/** What a column's "By color" pop-ups offer: the cell fills, font colours and conditional
 *  formatting icons its data rows show, each once, in row order. */
export interface ColumnColorChoices {
  fills: string[];
  fonts: string[];
  icons: ConditionalIconRef[];
}

/** Scan the data rows under `range`'s header in column `col`. Colours are lower-cased. */
export function columnColorChoices(state: State, range: Range, col: number): ColumnColorChoices {
  const fills = new Set<string>();
  const fonts = new Set<string>();
  const icons = new Map<string, ConditionalIconRef>();
  const overlays = evaluateConditional(state);
  for (let row = range.r0 + 1; row <= range.r1; row += 1) {
    const key = addrKey({ sheet: range.sheet, row, col });
    const fmt = state.format.formats.get(key);
    if (fmt?.fill) fills.add(fmt.fill.trim().toLocaleLowerCase());
    if (fmt?.color) fonts.add(fmt.color.trim().toLocaleLowerCase());
    const overlay = overlays.get(key);
    if (overlay?.iconKind !== undefined && overlay.iconSlot !== undefined) {
      icons.set(`${overlay.iconKind}:${overlay.iconSlot}`, {
        set: overlay.iconKind,
        slot: overlay.iconSlot,
      });
    }
  }
  return { fills: [...fills], fonts: [...fonts], icons: [...icons.values()] };
}
