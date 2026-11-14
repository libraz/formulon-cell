import { addrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import { evaluateConditional } from '../render/conditional.js';
import type { ConditionalIconRef } from '../store/types/conditional-rule.js';
import type { State } from '../store/types.js';
import { cellColorKey, NO_COLOR } from './filter.js';

/** What a column's "By color" pop-ups offer: the cell fills, font colours and conditional
 *  formatting icons its data rows show, each once, in row order, and whether a row has no fill
 *  (`noFill`) or no font colour of its own (`automaticFont`). */
export interface ColumnColorChoices {
  fills: string[];
  fonts: string[];
  icons: ConditionalIconRef[];
  noFill: boolean;
  automaticFont: boolean;
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
    fills.add(cellColorKey(fmt, 'cellColor'));
    fonts.add(cellColorKey(fmt, 'fontColor'));
    const overlay = overlays.get(key);
    if (overlay?.iconKind !== undefined && overlay.iconSlot !== undefined) {
      icons.set(`${overlay.iconKind}:${overlay.iconSlot}`, {
        set: overlay.iconKind,
        slot: overlay.iconSlot,
      });
    }
  }
  const listed = (colors: Set<string>): string[] => [...colors].filter((c) => c !== NO_COLOR);
  return {
    fills: listed(fills),
    fonts: listed(fonts),
    icons: [...icons.values()],
    noFill: fills.has(NO_COLOR),
    automaticFont: fonts.has(NO_COLOR),
  };
}
