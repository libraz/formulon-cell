/** A sheet edit that renumbers the sheets after it. */
export type SheetStructureEvent =
  | { kind: 'sheet-remove'; index: number }
  | { kind: 'sheet-move'; from: number; to: number }
  | { kind: 'sheet-copy'; index: number };

/** A sheet's index after a sheet edit; null for a removed sheet. */
export type SheetIndexMapper = (index: number) => number | null;

export const mapIndexForStructureEvent = (event: SheetStructureEvent): SheetIndexMapper => {
  // A copy is inserted at its index; every sheet from there on moves up one place.
  if (event.kind === 'sheet-copy') return (index) => (index >= event.index ? index + 1 : index);
  if (event.kind === 'sheet-remove') {
    return (index) => {
      if (index === event.index) return null;
      return index > event.index ? index - 1 : index;
    };
  }
  return (index) => {
    if (index === event.from) return event.to;
    if (event.from < event.to && index > event.from && index <= event.to) return index - 1;
    if (event.from > event.to && index >= event.to && index < event.from) return index + 1;
    return index;
  };
};
