import type { FormulaEditArrow } from './formula-edit-navigation.js';

export interface FormulaReferenceArrow {
  readonly direction: FormulaEditArrow;
  readonly extend: boolean;
  readonly jump: boolean;
}

/** Resolve the supported arrow modifiers shared by both formula edit surfaces. */
export function classifyFormulaReferenceArrow(
  event: KeyboardEvent,
  isMacPlatform: boolean,
): FormulaReferenceArrow | null {
  if (
    event.altKey ||
    event.isComposing ||
    event.key === 'Process' ||
    (event.ctrlKey && event.metaKey) ||
    (event.metaKey && !isMacPlatform)
  )
    return null;

  const direction =
    event.key === 'ArrowUp'
      ? 'up'
      : event.key === 'ArrowDown'
        ? 'down'
        : event.key === 'ArrowLeft'
          ? 'left'
          : event.key === 'ArrowRight'
            ? 'right'
            : null;
  if (!direction) return null;

  return {
    direction,
    extend: event.shiftKey,
    jump: event.ctrlKey || event.metaKey,
  };
}
