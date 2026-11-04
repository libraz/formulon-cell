export const finite = (value: number, fallback: number): number =>
  Number.isFinite(value) ? value : fallback;

export const clamp = (value: number, min: number, max: number): number =>
  Math.min(max, Math.max(min, value));
