/** Width in px of a run of text in the editor's font. */
export type MeasureText = (text: string) => number;

/** The horizontal box an editor lays its text out in, in client px. */
export interface EditorTextBox {
  left: number;
  width: number;
  paddingLeft: number;
  paddingRight: number;
  align: 'left' | 'right' | 'center';
}

/** Client x where the editor's text begins, given its rendered width. */
export function textStartOffset(box: EditorTextBox, textWidth: number): number {
  const inner = box.width - box.paddingLeft - box.paddingRight;
  if (box.align === 'right') return box.left + box.width - box.paddingRight - textWidth;
  if (box.align === 'center') return box.left + box.paddingLeft + (inner - textWidth) / 2;
  return box.left + box.paddingLeft;
}

/** The caret index nearest `offsetX` px from the start of `text`, stepping by
 *  code point so the caret never lands inside a surrogate pair. */
export function caretIndexAt(text: string, offsetX: number, measure: MeasureText): number {
  let best = 0;
  let bestDistance = Math.abs(offsetX);
  let index = 0;
  for (const char of text) {
    index += char.length;
    const distance = Math.abs(measure(text.slice(0, index)) - offsetX);
    if (distance < bestDistance) {
      best = index;
      bestDistance = distance;
    }
  }
  return best;
}

const measureContext = (): CanvasRenderingContext2D | null => {
  try {
    return document.createElement('canvas').getContext('2d');
  } catch {
    return null;
  }
};

/**
 * The caret index under `clientX` in a single-line editor, or null when the
 * text cannot be measured here (no canvas) or spans several lines.
 */
export function caretIndexAtClientX(
  input: HTMLTextAreaElement | HTMLInputElement,
  clientX: number,
): number | null {
  const text = input.value;
  if (text.includes('\n')) return null;
  const ctx = measureContext();
  if (!ctx) return null;
  const style = getComputedStyle(input);
  ctx.font = style.font || `${style.fontSize} ${style.fontFamily}`;
  const measure: MeasureText = (run) => ctx.measureText(run).width;
  const rect = input.getBoundingClientRect();
  const px = (value: string): number => Number.parseFloat(value) || 0;
  const align =
    style.textAlign === 'right' || style.textAlign === 'end'
      ? 'right'
      : style.textAlign === 'center'
        ? 'center'
        : 'left';
  const start = textStartOffset(
    {
      left: rect.left + px(style.borderLeftWidth),
      width: rect.width - px(style.borderLeftWidth) - px(style.borderRightWidth),
      paddingLeft: px(style.paddingLeft),
      paddingRight: px(style.paddingRight),
      align,
    },
    measure(text),
  );
  return caretIndexAt(text, clientX - start + input.scrollLeft, measure);
}
