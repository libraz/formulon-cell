// Minimal per-paint context shared by the free chrome painters (gridlines,
// headers, freeze dividers, traces). Bundles the canvas state that the host
// `GridRenderer` would otherwise read off `this.ctx` / `this.dpr` etc.

export interface ChromePaintContext {
  ctx: CanvasRenderingContext2D;
  dpr: number;
  cssWidth: number;
  cssHeight: number;
}
