import { isColGroupCollapsed, isRowGroupCollapsed } from '../../commands/outline.js';
import type { Sparkline, State } from '../../store/store.js';
import type { ResolvedTheme } from '../../theme/resolve.js';
import { type AxisLayout, gridOriginX, gridOriginY, type Rect } from '../geometry.js';
export interface CheckboxHit {
  rect: Rect;
}

const CHECKBOX_SIZE = 14;

/** Paint a centered checkbox glyph inside `bounds`. Returns the hit rect
 *  for pointer routing. The checked state is rendered as a filled square
 *  with a white check; unchecked is an outlined square. Theme accent
 *  carries the active fill so the box matches the host palette. */
export function paintCheckbox(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  checked: boolean,
  theme: ResolvedTheme,
): CheckboxHit {
  const cx = bounds.x + bounds.w / 2;
  const cy = bounds.y + bounds.h / 2;
  const x = Math.round(cx - CHECKBOX_SIZE / 2);
  const y = Math.round(cy - CHECKBOX_SIZE / 2);
  ctx.save();
  if (checked) {
    ctx.fillStyle = theme.accent;
    ctx.fillRect(x, y, CHECKBOX_SIZE, CHECKBOX_SIZE);
    ctx.strokeStyle = '#ffffff';
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.moveTo(x + 3, y + CHECKBOX_SIZE / 2);
    ctx.lineTo(x + CHECKBOX_SIZE / 2 - 1, y + CHECKBOX_SIZE - 4);
    ctx.lineTo(x + CHECKBOX_SIZE - 3, y + 3);
    ctx.stroke();
  } else {
    ctx.fillStyle = theme.bg;
    ctx.fillRect(x, y, CHECKBOX_SIZE, CHECKBOX_SIZE);
    ctx.strokeStyle = theme.ruleStrong;
    ctx.lineWidth = 1;
    ctx.strokeRect(x + 0.5, y + 0.5, CHECKBOX_SIZE - 1, CHECKBOX_SIZE - 1);
  }
  ctx.restore();
  return { rect: { x, y, w: CHECKBOX_SIZE, h: CHECKBOX_SIZE } };
}

/** One slot per outline level. The bracket spine sits at slot center. */
export const OUTLINE_BRACKET_SLOT = 14;
const TOGGLE_SIZE = 11;

/** Paint a small +/- toggle button. (cx, cy) is the center. */
function paintToggle(
  ctx: CanvasRenderingContext2D,
  theme: ResolvedTheme,
  cx: number,
  cy: number,
  collapsed: boolean,
): Rect {
  const half = TOGGLE_SIZE / 2;
  const x = Math.round(cx - half);
  const y = Math.round(cy - half);
  ctx.save();
  ctx.fillStyle = theme.bg;
  ctx.fillRect(x, y, TOGGLE_SIZE, TOGGLE_SIZE);
  ctx.strokeStyle = theme.ruleStrong;
  ctx.lineWidth = 1;
  ctx.strokeRect(x + 0.5, y + 0.5, TOGGLE_SIZE - 1, TOGGLE_SIZE - 1);
  ctx.strokeStyle = theme.fg;
  ctx.lineWidth = 1;
  ctx.beginPath();
  // Horizontal stroke (always drawn)
  ctx.moveTo(x + 2.5, y + half + 0.5);
  ctx.lineTo(x + TOGGLE_SIZE - 2.5, y + half + 0.5);
  if (collapsed) {
    // Vertical stroke for `+`
    ctx.moveTo(x + half + 0.5, y + 2.5);
    ctx.lineTo(x + half + 0.5, y + TOGGLE_SIZE - 2.5);
  }
  ctx.stroke();
  ctx.restore();
  return { x, y, w: TOGGLE_SIZE, h: TOGGLE_SIZE };
}

export interface OutlineToggleHit {
  axis: 'row' | 'col';
  level: number;
  /** First/last index of the contiguous run this toggle controls. */
  i0: number;
  i1: number;
  rect: Rect;
}

/** Walk both gutters and paint brackets + toggles. Returns toggle hit-rects so
 *  the pointer layer can route clicks to collapse/expand commands. */
export function paintOutlineGutters(
  ctx: CanvasRenderingContext2D,
  state: State,
  theme: ResolvedTheme,
  cols: AxisLayout,
  rows: AxisLayout,
  cssWidth: number,
  cssHeight: number,
): OutlineToggleHit[] {
  const hits: OutlineToggleHit[] = [];
  const layout = state.layout;
  const ox = gridOriginX(layout);
  const oy = gridOriginY(layout);

  // ── Row gutter (left of row-number column).
  if (layout.outlineRowGutter > 0) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, 0, layout.outlineRowGutter, cssHeight);
    ctx.clip();
    let maxLvl = 0;
    for (const v of layout.outlineRows.values()) if (v > maxLvl) maxLvl = v;
    for (let lvl = 1; lvl <= maxLvl; lvl += 1) {
      const slotCx = (lvl - 1) * OUTLINE_BRACKET_SLOT + OUTLINE_BRACKET_SLOT / 2;
      let runStartIdx = -1;
      for (let i = 0; i <= rows.visible.length; i += 1) {
        const r = rows.visible[i];
        const inRun = r != null && (layout.outlineRows.get(r) ?? 0) >= lvl;
        if (inRun && runStartIdx < 0) runStartIdx = i;
        if (!inRun && runStartIdx >= 0) {
          const r0 = rows.visible[runStartIdx];
          const r1 = rows.visible[i - 1];
          if (r0 === undefined || r1 === undefined) {
            runStartIdx = -1;
            continue;
          }
          const top = oy + (rows.positionAt.get(r0) ?? 0);
          const bottomRow = rows.positionAt.get(r1) ?? 0;
          const bottomH = rows.sizeAt.get(r1) ?? 0;
          const bottom = oy + bottomRow + bottomH;
          ctx.strokeStyle = theme.rule;
          ctx.lineWidth = 1;
          ctx.beginPath();
          ctx.moveTo(slotCx + 0.5, top);
          ctx.lineTo(slotCx + 0.5, bottom);
          // Foot tick on the bottom (the "summary row" side).
          ctx.moveTo(slotCx + 0.5, bottom - 0.5);
          ctx.lineTo(slotCx + 5.5, bottom - 0.5);
          ctx.stroke();
          const collapsed = isRowGroupCollapsed(layout, r0, r1);
          // Toggle sits on the row that should remain visible — typically the
          // row just below the band on desktop default ("summary below"). If
          // there is no such row visible, render at the bottom edge of the run.
          const summaryRow = r1 + 1;
          const summaryY =
            rows.positionAt.has(summaryRow) && !layout.hiddenRows.has(summaryRow)
              ? oy + (rows.positionAt.get(summaryRow) ?? 0) + (rows.sizeAt.get(summaryRow) ?? 0) / 2
              : bottom + TOGGLE_SIZE / 2 + 2;
          const rect = paintToggle(ctx, theme, slotCx, summaryY, collapsed);
          hits.push({ axis: 'row', level: lvl, i0: r0, i1: r1, rect });
          runStartIdx = -1;
        }
      }
    }
    ctx.restore();
  }

  // ── Col gutter (above col-letter row).
  if (layout.outlineColGutter > 0) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, 0, cssWidth, layout.outlineColGutter);
    ctx.clip();
    let maxLvl = 0;
    for (const v of layout.outlineCols.values()) if (v > maxLvl) maxLvl = v;
    for (let lvl = 1; lvl <= maxLvl; lvl += 1) {
      const slotCy = (lvl - 1) * OUTLINE_BRACKET_SLOT + OUTLINE_BRACKET_SLOT / 2;
      let runStartIdx = -1;
      for (let i = 0; i <= cols.visible.length; i += 1) {
        const c = cols.visible[i];
        const inRun = c != null && (layout.outlineCols.get(c) ?? 0) >= lvl;
        if (inRun && runStartIdx < 0) runStartIdx = i;
        if (!inRun && runStartIdx >= 0) {
          const c0 = cols.visible[runStartIdx];
          const c1 = cols.visible[i - 1];
          if (c0 === undefined || c1 === undefined) {
            runStartIdx = -1;
            continue;
          }
          const left = ox + (cols.positionAt.get(c0) ?? 0);
          const rightCol = cols.positionAt.get(c1) ?? 0;
          const rightW = cols.sizeAt.get(c1) ?? 0;
          const right = ox + rightCol + rightW;
          ctx.strokeStyle = theme.rule;
          ctx.lineWidth = 1;
          ctx.beginPath();
          ctx.moveTo(left, slotCy + 0.5);
          ctx.lineTo(right, slotCy + 0.5);
          ctx.moveTo(right - 0.5, slotCy + 0.5);
          ctx.lineTo(right - 0.5, slotCy + 5.5);
          ctx.stroke();
          const collapsed = isColGroupCollapsed(layout, c0, c1);
          const summaryCol = c1 + 1;
          const summaryX =
            cols.positionAt.has(summaryCol) && !layout.hiddenCols.has(summaryCol)
              ? ox + (cols.positionAt.get(summaryCol) ?? 0) + (cols.sizeAt.get(summaryCol) ?? 0) / 2
              : right + TOGGLE_SIZE / 2 + 2;
          const rect = paintToggle(ctx, theme, summaryX, slotCy, collapsed);
          hits.push({ axis: 'col', level: lvl, i0: c0, i1: c1, rect });
          runStartIdx = -1;
        }
      }
    }
    ctx.restore();
  }

  return hits;
}

const DEFAULT_SPARKLINE_COLOR = '#0078d4';
const DEFAULT_NEGATIVE_COLOR = '#d24545';

/** Inline mini-chart drawn inside a single cell rect. `values` is the resolved
 *  numeric series for `spec.source`; non-numeric source cells are filtered by
 *  the caller before invoking. No-op when the series is empty. */
export function paintSparkline(
  ctx: CanvasRenderingContext2D,
  rect: Rect,
  spec: Sparkline,
  values: readonly number[],
): void {
  if (values.length === 0) return;
  const inset = 2;
  const x = rect.x + inset;
  const y = rect.y + inset;
  const w = Math.max(0, rect.w - inset * 2);
  const h = Math.max(0, rect.h - inset * 2);
  if (w <= 0 || h <= 0) return;
  const color = spec.color ?? DEFAULT_SPARKLINE_COLOR;
  const negColor = spec.negativeColor ?? DEFAULT_NEGATIVE_COLOR;

  ctx.save();
  ctx.beginPath();
  ctx.rect(x, y, w, h);
  ctx.clip();
  ctx.globalAlpha = 1;

  if (spec.kind === 'line') {
    let min = Number.POSITIVE_INFINITY;
    let max = Number.NEGATIVE_INFINITY;
    for (const v of values) {
      if (v < min) min = v;
      if (v > max) max = v;
    }
    const span = max - min || 1;
    const stepX = values.length > 1 ? w / (values.length - 1) : 0;
    ctx.strokeStyle = color;
    ctx.lineWidth = 1.5;
    ctx.lineJoin = 'round';
    ctx.beginPath();
    for (let i = 0; i < values.length; i += 1) {
      const v = values[i] ?? 0;
      const px = x + (values.length > 1 ? i * stepX : w / 2);
      const py = y + h - ((v - min) / span) * h;
      if (i === 0) ctx.moveTo(px, py);
      else ctx.lineTo(px, py);
    }
    ctx.stroke();
  } else if (spec.kind === 'column') {
    let min = Number.POSITIVE_INFINITY;
    let max = Number.NEGATIVE_INFINITY;
    for (const v of values) {
      if (v < min) min = v;
      if (v > max) max = v;
    }
    // Anchor the baseline at zero when the range straddles it; otherwise rest
    // bars on the value extreme closest to zero so all bars stay positive-side.
    const baseline = min < 0 && max > 0 ? 0 : min >= 0 ? min : max;
    const span = Math.max(Math.abs(max - baseline), Math.abs(min - baseline)) || 1;
    const slot = w / values.length;
    const gap = Math.min(1, slot * 0.2);
    const barW = Math.max(1, slot - gap);
    const baseY = y + h - ((baseline - min) / Math.max(max - min, 1e-9)) * h;
    for (let i = 0; i < values.length; i += 1) {
      const v = values[i] ?? 0;
      const isNeg = v < 0;
      const fill = spec.showNegative && isNeg ? negColor : color;
      const px = x + i * slot;
      const barH = (Math.abs(v - baseline) / span) * h;
      const py = isNeg ? baseY : baseY - barH;
      ctx.fillStyle = fill;
      ctx.fillRect(px, py, barW, Math.max(1, barH));
    }
  } else {
    // win-loss
    const half = h / 2;
    const slot = w / values.length;
    const gap = Math.min(1, slot * 0.2);
    const barW = Math.max(1, slot - gap);
    const midY = y + half;
    for (let i = 0; i < values.length; i += 1) {
      const v = values[i] ?? 0;
      if (v === 0) continue;
      const isNeg = v < 0;
      const fill = spec.showNegative && isNeg ? negColor : color;
      ctx.fillStyle = fill;
      const px = x + i * slot;
      if (isNeg) ctx.fillRect(px, midY, barW, Math.max(1, half - 1));
      else ctx.fillRect(px, midY - Math.max(1, half - 1), barW, Math.max(1, half - 1));
    }
  }

  ctx.restore();
}
