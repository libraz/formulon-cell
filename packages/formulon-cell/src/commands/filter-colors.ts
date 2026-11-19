import { addrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import { type ConditionalCellOverlay, evaluateConditional } from '../render/conditional.js';
import type { FilterColorSpec } from '../store/store.js';
import type { ConditionalIconRef } from '../store/types/conditional-rule.js';
import type { State } from '../store/types.js';

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

export const normalizeColorKey = (color: string): string => color.trim().toLocaleLowerCase();

/** The colour key of a cell with no fill (No Fill) or no font colour of its own (Automatic). */
export const NO_COLOR = 'none';

/** The key a colour filter or sort level matches in a cell: its fill or font colour, lower-cased, or
 *  `NO_COLOR` when it has none. */
export const cellColorKey = (
  fmt: { fill?: string; color?: string } | undefined,
  kind: 'cellColor' | 'fontColor',
): string => {
  const actual = kind === 'cellColor' ? fmt?.fill : fmt?.color;
  return actual ? normalizeColorKey(actual) : NO_COLOR;
};

export function colorMatches(
  state: State,
  sheet: number,
  row: number,
  col: number,
  color: FilterColorSpec,
  icons: ReadonlyMap<string, ConditionalCellOverlay> | null,
): boolean {
  const key = addrKey({ sheet, row, col });
  if (color.kind === 'cellIcon') {
    const overlay = icons?.get(key);
    return overlay?.iconKind === color.icon.set && overlay.iconSlot === color.icon.slot;
  }
  return cellColorKey(state.format.formats.get(key), color.kind) === color.color;
}
