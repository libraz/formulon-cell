/** Keyboard model for the demo's select-only combobox (the panel's language
 * picker). The React and Vue components own the open/active state and apply
 * the action this resolves, so both hosts share one interaction contract. */

export interface DemoSelectOption {
  value: string;
  label: string;
}

export type DemoSelectKeyAction =
  | { kind: 'open'; index: number }
  | { kind: 'move'; index: number }
  | { kind: 'commit'; index: number; keepDefault: boolean }
  | { kind: 'close' };

/** Resolve a keydown on the combobox. `activeIndex` is the highlighted option
 *  while open; `selectedIndex` seeds the highlight on open. Returns null for
 *  keys the combobox leaves alone. */
export const resolveDemoSelectKey = (
  key: string,
  open: boolean,
  activeIndex: number,
  selectedIndex: number,
  count: number,
): DemoSelectKeyAction | null => {
  if (count === 0) return null;
  const last = count - 1;
  if (!open) {
    switch (key) {
      case 'ArrowDown':
      case 'ArrowUp':
      case 'Enter':
      case ' ':
        return { kind: 'open', index: Math.max(0, selectedIndex) };
      case 'Home':
        return { kind: 'open', index: 0 };
      case 'End':
        return { kind: 'open', index: last };
      default:
        return null;
    }
  }
  switch (key) {
    case 'ArrowDown':
      return { kind: 'move', index: Math.min(last, activeIndex + 1) };
    case 'ArrowUp':
      return { kind: 'move', index: Math.max(0, activeIndex - 1) };
    case 'Home':
      return { kind: 'move', index: 0 };
    case 'End':
      return { kind: 'move', index: last };
    case 'Enter':
    case ' ':
      return { kind: 'commit', index: activeIndex, keepDefault: false };
    case 'Tab':
      return { kind: 'commit', index: activeIndex, keepDefault: true };
    case 'Escape':
      return { kind: 'close' };
    default:
      return null;
  }
};
