import type { CellFormat } from './types.js';

const SOLID_OR_PATTERN_FIELDS = ['fill', 'fillPattern', 'fillPatternColor'] as const;

/** Apply the fill-kind rule to `next`, the result of spreading `patch` over a
 *  base: a solid, pattern or empty fill in `patch` drops a gradient, and a
 *  gradient in `patch` drops the solid and pattern fields. */
export function withFillReplacement<T extends Partial<CellFormat>>(
  next: T,
  patch: Partial<CellFormat>,
): T {
  if (patch.fillGradient !== undefined) {
    for (const field of SOLID_OR_PATTERN_FIELDS) delete next[field];
  } else if (
    !Object.hasOwn(patch, 'fillGradient') &&
    SOLID_OR_PATTERN_FIELDS.some((field) => Object.hasOwn(patch, field))
  ) {
    delete next.fillGradient;
  }
  return next;
}

/** `patch` merged over `base`: borders merge side by side and the fill-kind
 *  rule of `withFillReplacement` applies. */
export function mergeFormatPatch(
  base: CellFormat | undefined,
  patch: Partial<CellFormat>,
): CellFormat {
  const next: CellFormat = { ...(base ?? {}), ...patch };
  if (patch.borders) next.borders = { ...(base?.borders ?? {}), ...patch.borders };
  return withFillReplacement(next, patch);
}
