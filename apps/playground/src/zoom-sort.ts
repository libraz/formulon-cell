import { type SpreadsheetInstance, setSheetZoom } from '@libraz/formulon-cell';

export function setupZoomControls(getInst: () => SpreadsheetInstance | null): {
  refreshZoom(): void;
} {
  const zoomDisplay = document.getElementById('zoom-display');
  const zoomRailFill = document.getElementById('zoom-rail-fill');
  const zoomRailThumb = document.getElementById('zoom-rail-thumb');
  const zMin = 0.25;
  const zMax = 3.0;
  const refreshZoom = (): void => {
    const inst = getInst();
    if (!inst) return;
    const z = inst.store.getState().viewport.zoom;
    if (zoomDisplay) zoomDisplay.textContent = `${Math.round(z * 100)}%`;
    const pct = Math.max(0, Math.min(1, (z - zMin) / (zMax - zMin))) * 100;
    if (zoomRailFill) zoomRailFill.style.width = `${pct}%`;
    if (zoomRailThumb) zoomRailThumb.style.left = `${pct}%`;
  };
  const stepZoom = (delta: number): void => {
    const inst = getInst();
    if (!inst) return;
    const z = inst.store.getState().viewport.zoom;
    const next = Math.max(zMin, Math.min(zMax, Math.round((z + delta) * 100) / 100));
    if (next === z) return;
    setSheetZoom(inst.store, next, inst.workbook);
    refreshZoom();
  };
  zoomDisplay?.addEventListener('click', () => {
    const inst = getInst();
    if (!inst) return;
    const z = inst.store.getState().viewport.zoom;
    const next = z >= 1.5 ? 0.75 : Math.round((z + 0.25) * 100) / 100;
    setSheetZoom(inst.store, next, inst.workbook);
    refreshZoom();
  });
  document.getElementById('btn-zoom-out')?.addEventListener('click', () => stepZoom(-0.1));
  document.getElementById('btn-zoom-in')?.addEventListener('click', () => stepZoom(0.1));
  return { refreshZoom };
}
