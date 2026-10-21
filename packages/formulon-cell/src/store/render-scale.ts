import type { LayoutSlice, ViewportSlice } from './types.js';

/**
 * Canonical CSS px → screen px factor: viewport zoom times the platform
 * screen scale. A zero or non-finite zoom counts as 1; without a viewport
 * the result is the screen scale alone.
 */
export function renderScale(
  layout: Pick<LayoutSlice, 'screenScale'>,
  viewport?: Pick<ViewportSlice, 'zoom'>,
): number {
  const zoom = viewport?.zoom;
  const z = zoom !== undefined && Number.isFinite(zoom) && zoom > 0 ? zoom : 1;
  return z * (layout.screenScale ?? 1);
}
