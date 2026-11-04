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

/** Sheet zoom range as a multiplier: the desktop spreadsheet's 10%..400%. */
export const ZOOM_MIN = 0.1;
export const ZOOM_MAX = 4;

export const clampZoom = (zoom: number): number => Math.max(ZOOM_MIN, Math.min(ZOOM_MAX, zoom));

/** Zoom in/out buttons step to the next multiple of 10%, so 103% goes to
 *  110% and back to 100%. */
export function steppedZoom(zoom: number, direction: 1 | -1): number {
  const pct = Math.round(zoom * 100);
  const next = direction > 0 ? Math.floor(pct / 10) * 10 + 10 : Math.ceil(pct / 10) * 10 - 10;
  return clampZoom(next / 100);
}

/** Status-bar slider positions run 0..ZOOM_SLIDER_MAX like the desktop
 *  slider; one increment of ZOOM_SLIDER_STEP is 9% below 100% and 30% above. */
export const ZOOM_SLIDER_MAX = 2000;
export const ZOOM_SLIDER_STEP = 100;
const ZOOM_SLIDER_MID = ZOOM_SLIDER_MAX / 2;

/** The status-bar slider keeps 100% at its centre: the left half spans
 *  10%..100% and the right half 100%..400%. */
export function zoomToSliderPosition(zoom: number): number {
  const z = clampZoom(zoom);
  return z <= 1
    ? ((z - ZOOM_MIN) / (1 - ZOOM_MIN)) * ZOOM_SLIDER_MID
    : ZOOM_SLIDER_MID + ((z - 1) / (ZOOM_MAX - 1)) * ZOOM_SLIDER_MID;
}

export function sliderPositionToZoom(position: number): number {
  const p = Math.max(0, Math.min(ZOOM_SLIDER_MAX, position));
  const zoom =
    p <= ZOOM_SLIDER_MID
      ? ZOOM_MIN + (p / ZOOM_SLIDER_MID) * (1 - ZOOM_MIN)
      : 1 + ((p - ZOOM_SLIDER_MID) / ZOOM_SLIDER_MID) * (ZOOM_MAX - 1);
  return clampZoom(Math.round(zoom * 100) / 100);
}
