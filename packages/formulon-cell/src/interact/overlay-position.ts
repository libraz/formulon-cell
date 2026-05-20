export interface ViewportPanelPositionOptions {
  pad?: number;
  fallbackWidth?: number;
  fallbackHeight?: number;
}

export interface ViewportPanelPosition {
  x: number;
  y: number;
}

export const viewportSize = (): { width: number; height: number } => ({
  width: window.innerWidth || document.documentElement.clientWidth || 1024,
  height: window.innerHeight || document.documentElement.clientHeight || 768,
});

export const clamp = (value: number, min: number, max: number): number =>
  Math.min(Math.max(value, min), Math.max(min, max));

export const panelSize = (
  panel: HTMLElement,
  fallbackWidth = 0,
  fallbackHeight = 0,
): { width: number; height: number } => {
  const rect = panel.getBoundingClientRect();
  return {
    width: Math.ceil(rect.width || panel.offsetWidth || fallbackWidth),
    height: Math.ceil(rect.height || panel.offsetHeight || fallbackHeight),
  };
};

export const clampPanelToViewport = (
  panel: HTMLElement,
  x: number,
  y: number,
  options: ViewportPanelPositionOptions = {},
): ViewportPanelPosition => {
  const pad = options.pad ?? 4;
  const { width, height } = panelSize(panel, options.fallbackWidth, options.fallbackHeight);
  const viewport = viewportSize();
  return {
    x: clamp(x, pad, viewport.width - width - pad),
    y: clamp(y, pad, viewport.height - height - pad),
  };
};
