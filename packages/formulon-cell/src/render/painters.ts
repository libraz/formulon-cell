// Barrel for the render painters. The actual implementations are split by
// responsibility under ./painters/:
//   - types.ts    — shared CellPaintCtx / TextMetricsBox / TextVAlign
//   - trace.ts    — formula-trace dots, arrows, ref-highlight overlay
//   - markers.ts  — error/validation/lock/comment markers + chevrons
//   - cell.ts     — cell background/fill/borders + conditional-format icon set
//   - text.ts     — cell text rendering, font/align/wrap/shrink helpers
//   - handles.ts  — active outline, fill handle, marquee, spill outlines
//   - controls.ts — checkbox + sparkline painters (pre-existing split)

export * from './painters/cell.js';
export * from './painters/controls.js';
export * from './painters/handles.js';
export * from './painters/markers.js';
export * from './painters/text.js';
export * from './painters/trace.js';
export * from './painters/types.js';
