import { isHeaderRow, tableForCell } from '../../commands/format-as-table.js';
import { formatA1FormulaAsR1C1 } from '../../commands/refs.js';
import { addrKey } from '../../engine/address.js';
import { effectiveCellFormatFromEngine } from '../../engine/cell-format-sync.js';
import { evaluateCfFromEngine } from '../../engine/cf-sync.js';
import { formulaForEdit } from '../../engine/edit-seed.js';
import { numFmtToFormatCode } from '../../engine/format-writeback.js';
import { makeRangeResolver, type RangeResolver } from '../../engine/range-resolver.js';
import type { CellValue } from '../../engine/types.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import { numberFormatPaint } from '../../format/number-format.js';
import { rangeValidationAt } from '../../store/range-validations.js';
import { renderScale } from '../../store/render-scale.js';
import { rangeContainsAddr } from '../../store/selection-geometry.js';
import type { CellFormat } from '../../store/store.js';
import type { ResolvedTheme } from '../../theme/resolve.js';
import { evaluateConditional } from '../conditional.js';
import {
  type AxisLayout,
  cellRectIn,
  gridOriginX,
  gridOriginY,
  type Rect,
  rangeRects,
  type ViewState,
} from '../geometry.js';
import {
  CONDITIONAL_ICON_GUTTER,
  paintCellBackground,
  paintCellBorders,
  paintCellFill,
  paintCellText,
  paintCommentMarker,
  paintConditionalIcon,
  paintErrorTriangle,
  paintLockMarker,
  paintTableHeaderChevron,
  paintValidationCircle,
  paintValidationTriangle,
} from '../painters.js';
import { paintCellSparkline } from '../sparkline.js';
import { paintDataBar } from './data-bar.js';
import {
  detectErrorKind,
  detectValidationViolation,
  ERROR_TRIANGLE_COLOR,
  type ErrorTriangleHit,
  isPlainTextOverflowCandidate,
  normalizeFormatLocale,
  setErrorTriangleHits,
  VALIDATION_TRIANGLE_COLOR,
} from './hit-state.js';
import { connectedRects } from './range-rects.js';
import { tableCellFormat } from './table-format.js';

export type CellDisplayResolver = (
  addr: { sheet: number; row: number; col: number },
  value: CellValue,
  formula: string | null,
  format: CellFormat | undefined,
) => string | null;

export interface CellPaintContext {
  ctx: CanvasRenderingContext2D;
  wb: WorkbookHandle | null;
  locale: string;
  getDisplay: CellDisplayResolver | undefined;
  /** The active sheet's background image once loaded, otherwise null. */
  sheetBackgroundImage: HTMLImageElement | null;
  /** The host renders the Mac profile, which draws Outline and Shadow. */
  fontEffects: boolean;
}

export function paintCells(
  cellCtx: CellPaintContext,
  state: ViewState,
  theme: ResolvedTheme,
  cols: AxisLayout,
  rows: AxisLayout,
): void {
  const ctx = cellCtx.ctx;
  const {
    layout,
    data,
    selection,
    format,
    merges,
    sparkline,
    errorIndicators,
    protection,
    tables,
  } = state;
  const rtl = layout.rtl;
  const active = selection.active;
  const wb = cellCtx.wb;
  const visibleBounds = {
    r0: state.viewport.rowStart,
    c0: state.viewport.colStart,
    r1: Math.max(state.viewport.rowStart, state.viewport.rowStart + state.viewport.rowCount - 1),
    c1: Math.max(state.viewport.colStart, state.viewport.colStart + state.viewport.colCount - 1),
  };
  const conditional = new Map(evaluateConditional(state, wb, cellCtx.locale, visibleBounds));
  const sparklines = sparkline.sparklines;
  const ignored = errorIndicators.ignoredErrors;
  const validationCircles = errorIndicators.validationCircles;
  // Lock-icon overlay only fires when the active sheet is currently
  // protected. Pre-compute the flag so the per-cell loop can skip the
  // Map lookup on every iteration.
  const sheetProtected = protection.protectedSheets.has(data.sheetIndex);
  const threadedRootKeys = new Set(
    (state.threadedComments?.bySheet.get(data.sheetIndex) ?? [])
      .filter((comment) => comment.parentId.length === 0)
      .map((comment) => addrKey({ sheet: data.sheetIndex, row: comment.row, col: comment.col })),
  );
  // RangeResolver is only needed for list-validation re-resolution. We
  // build it lazily — most viewport repaints don't have a single DV cell,
  // so allocating the resolver up-front would be wasted work.
  const wbForResolver = cellCtx.wb;
  let resolver: RangeResolver | undefined;
  const getResolver = (): RangeResolver | undefined => {
    if (resolver) return resolver;
    if (!wbForResolver) return undefined;
    resolver = makeRangeResolver(wbForResolver, data.sheetIndex);
    return resolver;
  };
  const triangleHits: ErrorTriangleHit[] = [];
  const locale = normalizeFormatLocale(cellCtx.locale);
  // Engine-side CF (rules loaded from .xlsx). Merged on top — engine rules
  // currently win per field for cells with overlap. Restricted to the
  // visible viewport rect so we don't pay for off-screen cells.
  const effectiveFormats = new Map<string, Partial<CellFormat> | null>();
  const effectiveFormatAt = (row: number, col: number): Partial<CellFormat> | undefined => {
    if (!wb?.capabilities.effectiveStyle || typeof wb.getEffectiveStyle !== 'function') {
      return undefined;
    }
    return (
      effectiveCellFormatFromEngine(wb, data.sheetIndex, row, col, effectiveFormats) ?? undefined
    );
  };
  if (wb?.capabilities.conditionalFormat) {
    const sheet = state.data.sheetIndex;
    const engineCf = evaluateCfFromEngine(
      wb,
      sheet,
      visibleBounds.r0,
      visibleBounds.c0,
      visibleBounds.r1,
      visibleBounds.c1,
    );
    for (const [k, v] of engineCf) {
      const merged = { ...(conditional.get(k) ?? {}), ...v };
      conditional.set(k, merged);
    }
  }

  const rng = selection.range;
  const objectSelected = state.ui.selectedObject !== null;
  if (!objectSelected && (rng.r0 !== rng.r1 || rng.c0 !== rng.c1)) {
    ctx.fillStyle = theme.accentSoft;
    const rects = rangeRects(layout, state.viewport, rng);
    for (const r of rects) ctx.fillRect(r.x, r.y, r.w, r.h);
  }
  // Disjoint multi-range selection (Ctrl/Cmd+click). Paint each extra band
  // with the same accent so the user can read all members at a glance.
  const extras = selection.extraRanges;
  if (!objectSelected && extras && extras.length > 0) {
    ctx.fillStyle = theme.accentSoft;
    for (const er of extras) {
      if (er.r0 > er.r1 || er.c0 > er.c1) continue;
      for (const r of rangeRects(layout, state.viewport, er)) {
        ctx.fillRect(r.x, r.y, r.w, r.h);
      }
    }
  }

  // Compute merged cell bounds from the same freeze/RTL-aware projection
  // used by selection outlines. The old local width/height summation mixed
  // logical and screen coordinates when a merge crossed a frozen boundary.
  const mergeRectsAt = (row: number, col: number): Rect[] => {
    const range = merges.byAnchor.get(addrKey({ sheet: data.sheetIndex, row, col }));
    return range ? rangeRects(layout, state.viewport, range) : [];
  };
  const mergeBounds = (row: number, col: number): Rect | null =>
    connectedRects(mergeRectsAt(row, col))[0] ?? null;

  // The sheet background is painted before grid lines. Repaint it inside a
  // merge before its content so internal gridlines cannot show through. If
  // the background image is already loaded, restore the same repeat pattern
  // instead of replacing it with an opaque theme fill.
  const paintMergeBase = (bounds: Rect): void => {
    const image = cellCtx.sheetBackgroundImage;
    const pattern =
      image &&
      image.naturalWidth > 0 &&
      image.naturalHeight > 0 &&
      typeof ctx.createPattern === 'function'
        ? ctx.createPattern(image, 'repeat')
        : null;
    if (pattern) {
      const ox = layout.rtl ? 0 : gridOriginX(layout);
      const oy = gridOriginY(layout);
      ctx.save();
      ctx.translate(ox, oy);
      ctx.fillStyle = pattern;
      ctx.fillRect(bounds.x - ox, bounds.y - oy, bounds.w, bounds.h);
      ctx.restore();
      return;
    }
    ctx.fillStyle = theme.bg;
    ctx.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);
  };
  const paintMergeSurface = (
    bounds: Rect,
    isActive: boolean,
    isSelected: boolean | undefined,
  ): void => {
    paintMergeBase(bounds);
    if (isActive) {
      paintCellBackground({
        ctx,
        theme,
        bounds,
        value: { kind: 'blank' },
        formula: null,
        isActive: true,
        isInRange: false,
      });
    } else if (isSelected) {
      paintCellBackground({
        ctx,
        theme,
        bounds,
        value: { kind: 'blank' },
        formula: null,
        isActive: false,
        isInRange: true,
      });
    }
  };
  // Overflow and centre-across both spill into the following columns, which
  // the mirror puts to the left of the anchor on a right-to-left sheet — so
  // the widened rect has to move its left edge as well as grow.
  const spillRect = (base: Rect, extra: number): Rect =>
    layout.rtl ? { ...base, x: base.x - extra, w: base.w + extra } : { ...base, w: base.w + extra };
  const overflowBounds = (row: number, col: number, base: Rect): Rect => {
    let w = 0;
    for (let nextCol = col + 1; cols.positionAt.has(nextCol); nextCol += 1) {
      const nextKey = addrKey({ sheet: data.sheetIndex, row, col: nextCol });
      const nextCell = data.cells.get(nextKey);
      const nextFmt = format.formats.get(nextKey);
      const nextTable = tableForCell(tables.tables, data.sheetIndex, row, nextCol);
      if (
        merges.byCell.has(nextKey) ||
        merges.byAnchor.has(nextKey) ||
        nextTable ||
        nextFmt?.fill ||
        nextFmt?.borders ||
        (nextCell && (nextCell.formula || nextCell.value.kind !== 'blank'))
      ) {
        break;
      }
      w += cols.sizeAt.get(nextCol) ?? 0;
    }
    return w === 0 ? base : spillRect(base, w);
  };
  const centerContinuousBounds = (row: number, col: number, base: Rect): Rect => {
    let extra = 0;
    for (let nextCol = col + 1; cols.positionAt.has(nextCol); nextCol += 1) {
      const nextKey = addrKey({ sheet: data.sheetIndex, row, col: nextCol });
      const nextCell = data.cells.get(nextKey);
      const nextFmt = format.formats.get(nextKey);
      const nextTable = tableForCell(tables.tables, data.sheetIndex, row, nextCol);
      if (
        nextTable ||
        merges.byCell.has(nextKey) ||
        merges.byAnchor.has(nextKey) ||
        nextFmt?.align !== 'centerContinuous' ||
        (nextCell && (nextCell.formula || nextCell.value.kind !== 'blank'))
      ) {
        break;
      }
      extra += cols.sizeAt.get(nextCol) ?? 0;
    }
    return extra === 0 ? base : spillRect(base, extra);
  };

  // The anchor can be scrolled or hidden while a visible body segment of the
  // merge remains on screen. Clear that segment here; the normal anchor walk
  // below handles the common visible-anchor case and then paints content.
  for (const [anchorKey, merge] of merges.byAnchor) {
    if (merge.sheet !== data.sheetIndex) continue;
    const rects = rangeRects(layout, state.viewport, merge);
    const anchorVisible = rows.positionAt.has(merge.r0) && cols.positionAt.has(merge.c0);
    if (rects.length === 0 || anchorVisible) continue;
    const anchor = { sheet: merge.sheet, row: merge.r0, col: merge.c0 };
    const selected =
      rangeContainsAddr(rng, anchor) ||
      extras?.some((extra) => rangeContainsAddr(extra, anchor)) === true;
    const isActive =
      active.sheet === merge.sheet && active.row === merge.r0 && active.col === merge.c0;
    const anchorFormat = format.formats.get(anchorKey);
    for (const rect of rects) {
      paintMergeSurface(rect, isActive, selected);
      if (anchorFormat?.fill || anchorFormat?.fillPattern || anchorFormat?.fillGradient) {
        paintCellFill({
          ctx,
          theme,
          bounds: rect,
          value: { kind: 'blank' },
          formula: null,
          isActive,
          isInRange: selected,
          format: anchorFormat,
        });
      }
    }
  }

  // Single visible-cells walk paints both static format fills (for blank
  // formatted cells too) and cell content. Iterating format.formats here
  // would be O(formats), which dominates on sheets with thousands of
  // formatted cells; a viewport-sized grid is bounded.
  for (const r of rows.visible) {
    for (const c of cols.visible) {
      const key = addrKey({ sheet: data.sheetIndex, row: r, col: c });
      // Skip cells hidden inside a merge — only the anchor paints.
      if (merges.byCell.has(key)) continue;
      const cell = data.cells.get(key);
      const fmt = format.formats.get(key);
      // The engine may select a row/column XF for this cell even when the
      // store has no direct format entry. Keep this projection local to the
      // paint: materialising it into FormatSlice would turn inherited style
      // into a direct XF on the next flush.
      const engineFmt = effectiveFormatAt(r, c);
      const table = tableForCell(tables.tables, data.sheetIndex, r, c);
      const tableFmt = table ? tableCellFormat(table, r, c) : undefined;
      const isMergeAnchor = merges.byAnchor.has(key);
      // Use the union of contiguous visible rectangles for content bounds,
      // but retain every visible quadrant for the merged surface and static
      // format fill. A merge crossing a frozen row/column is projected into
      // more than one rect; painting only one leaves the other quadrants
      // with a different background or exposed gridline.
      const mergedRects = isMergeAnchor ? mergeRectsAt(r, c) : [];
      const spark = sparklines.get(key);
      const isActive = r === active.row && c === active.col;
      const hasValidationCircle = validationCircles.has(key);
      const hasThreadedComment = threadedRootKeys.has(key);
      // Render-worthy when there's data, a merge anchor, a static fill, or a
      // sparkline/table host — all reasons to paint into an otherwise blank cell.
      if (
        !cell &&
        !isActive &&
        !isMergeAnchor &&
        !engineFmt?.fill &&
        !engineFmt?.fillPattern &&
        !engineFmt?.fillGradient &&
        !engineFmt?.borders &&
        !fmt?.fill &&
        !fmt?.fillPattern &&
        !fmt?.fillGradient &&
        !fmt?.borders &&
        !tableFmt?.fill &&
        !tableFmt?.fillPattern &&
        !tableFmt?.borders &&
        !spark &&
        !hasValidationCircle &&
        !hasThreadedComment
      )
        continue;
      const bounds = mergeBounds(r, c) ?? cellRectIn(layout, cols, rows, r, c);
      const isInRange = rangeContainsAddr(rng, { sheet: data.sheetIndex, row: r, col: c });
      const isInExtraRange = extras?.some((extra) =>
        rangeContainsAddr(extra, { sheet: data.sheetIndex, row: r, col: c }),
      );

      const overlay = conditional.get(key);
      const baseFmt: typeof fmt =
        engineFmt || tableFmt || fmt
          ? {
              ...engineFmt,
              ...tableFmt,
              ...fmt,
            }
          : undefined;
      const effectiveFmt: typeof fmt =
        overlay && (overlay.fill || overlay.color || overlay.bold || overlay.italic)
          ? {
              ...baseFmt,
              fill: overlay.fill ?? baseFmt?.fill,
              fillGradient: overlay.fill ? undefined : baseFmt?.fillGradient,
              color: overlay.color ?? baseFmt?.color,
              bold: overlay.bold || baseFmt?.bold,
              italic: overlay.italic || baseFmt?.italic,
              underline: overlay.underline || baseFmt?.underline,
              strike: overlay.strike || baseFmt?.strike,
            }
          : baseFmt;

      if (isMergeAnchor) {
        for (const mergedRect of mergedRects) {
          paintMergeSurface(mergedRect, isActive, isInRange || isInExtraRange);
        }
      }

      const value: CellValue = cell?.value ?? { kind: 'blank' };
      const storedFormula = cell?.formula ?? null;
      const cellAddr = { sheet: data.sheetIndex, row: r, col: c };
      let formula = storedFormula;
      if (storedFormula && state.ui.showFormulas === true) {
        if (wb) {
          try {
            formula = formulaForEdit(wb, cellAddr, storedFormula, state.ui.r1c1) ?? storedFormula;
          } catch {
            // A detached or older workbook can expose only part of the edit
            // seed API. Preserve the standalone renderer's conversion below.
            formula = state.ui.r1c1
              ? formatA1FormulaAsR1C1(storedFormula, cellAddr)
              : storedFormula;
          }
        }
        if (!wb && state.ui.r1c1 === true) {
          formula = formatA1FormulaAsR1C1(storedFormula, { row: r, col: c });
        }
      }
      const customDisplay =
        cellCtx.getDisplay?.(
          { sheet: data.sheetIndex, row: r, col: c },
          value,
          storedFormula,
          fmt,
        ) ?? null;
      let nativeDisplay: string | null = null;
      const formatCode = fmt?.numFmt ? numFmtToFormatCode(fmt.numFmt, wb) : null;
      if (customDisplay === null && state.ui.showFormulas !== true && wb) {
        // A just-authored number format may be present in the store before a
        // mirror flush reaches the native XF table. Format that value directly
        // with formulon; loaded/inherited formats can use the workbook cell
        // renderer, which also applies its date system and effective XF.
        if (formatCode && !overlay && typeof wb.formatValueForDisplay === 'function') {
          nativeDisplay = wb.formatValueForDisplay(value, formatCode);
        } else if (
          !formatCode &&
          !overlay &&
          !fmt?.hyperlinkDisplay &&
          typeof wb.getDisplayText === 'function'
        ) {
          nativeDisplay = wb.getDisplayText(data.sheetIndex, r, c);
          // A newly entered value can reach the store one paint before the
          // engine mirror. An empty native result then means "not mirrored"
          // for a nonblank value; let the painter's deterministic formatter
          // handle that transient state.
          if (nativeDisplay === '' && value.kind !== 'blank') nativeDisplay = null;
        }
      }
      const displayOverride = customDisplay ?? nativeDisplay;
      const paintCtx = {
        ctx,
        theme,
        bounds,
        value,
        formula,
        isActive,
        isInRange,
        format: effectiveFmt,
        showFormulas: state.ui.showFormulas === true,
        showZeros: state.ui.showZeros !== false,
        displayOverride,
        locale,
        date1904: state.ui.date1904 === true,
        rtl,
        zoom: renderScale(state.layout, state.viewport),
        formatPaint:
          formatCode && customDisplay === null && value.kind === 'number'
            ? numberFormatPaint(formatCode, value.value)
            : null,
        fontEffects: cellCtx.fontEffects,
      };

      // Static format fill OR overlay fill — both painted via paintCellFill,
      // which reads `format.fill`. The merged effectiveFmt already prefers
      // overlay over static, so a single call yields the right result.
      if (effectiveFmt?.fill || effectiveFmt?.fillPattern || effectiveFmt?.fillGradient) {
        if (isMergeAnchor && mergedRects.length > 0) {
          for (const mergedRect of mergedRects) {
            paintCellFill({ ...paintCtx, bounds: mergedRect });
          }
        } else {
          paintCellFill(paintCtx);
        }
      }
      if (isActive && !effectiveFmt?.fill && !effectiveFmt?.fillGradient)
        paintCellBackground(paintCtx);
      if (overlay) paintDataBar(ctx, bounds, overlay);
      const hideConditionalValue =
        overlay?.showValue === false &&
        (overlay.bar !== undefined || (overlay.iconKind && overlay.iconSlot !== undefined));
      // Icon-set: paint the glyph in the leading gutter and inset the text by
      // the gutter width so the value reads cleanly next to the icon.
      const tableHeader = table ? isHeaderRow(table, r, c) : false;
      if (overlay?.iconKind && overlay.iconSlot !== undefined) {
        paintConditionalIcon(ctx, bounds, overlay.iconKind, overlay.iconSlot, rtl);
        if (!hideConditionalValue) {
          const insetBounds = {
            x: rtl ? bounds.x : bounds.x + CONDITIONAL_ICON_GUTTER,
            y: bounds.y,
            w: bounds.w - CONDITIONAL_ICON_GUTTER,
            h: bounds.h,
          };
          paintCellText({ ...paintCtx, bounds: insetBounds });
        }
      } else if (hideConditionalValue) {
        // Conditional formatting's "Show Bar/Icon Only" suppresses the
        // value text while still leaving fills, bars, and icons visible.
      } else if (tableHeader) {
        paintCellText({
          ...paintCtx,
          bounds: {
            ...bounds,
            x: rtl ? bounds.x + 20 : bounds.x,
            w: Math.max(0, bounds.w - 20),
          },
        });
      } else if (effectiveFmt?.align === 'centerContinuous' && !isMergeAnchor) {
        paintCellText({ ...paintCtx, bounds: centerContinuousBounds(r, c, bounds) });
      } else if (
        isPlainTextOverflowCandidate({
          value,
          formula,
          format: effectiveFmt,
          showFormulas: state.ui.showFormulas === true,
          displayOverride,
          tableHeader,
          hasIcon: false,
          isMergeAnchor,
        })
      ) {
        paintCellText({ ...paintCtx, bounds: overflowBounds(r, c, bounds) });
      } else {
        paintCellText(paintCtx);
      }
      if (tableHeader) paintTableHeaderChevron(ctx, bounds, theme, rtl);
      if (spark) paintCellSparkline(ctx, bounds, spark, state, cellCtx.wb);
      if (engineFmt?.borders && !fmt?.borders) {
        paintCellBorders({
          ctx,
          theme,
          bounds,
          value,
          formula,
          isActive,
          isInRange,
          format: effectiveFmt,
        });
      }
      if (fmt?.comment || hasThreadedComment) paintCommentMarker(ctx, bounds, rtl);
      if (hasValidationCircle) paintValidationCircle(ctx, bounds, VALIDATION_TRIANGLE_COLOR);
      // Lock-icon overlay — only when the sheet is protected AND the cell
      // is explicitly unlocked, signalling which cells the user can still
      // type into despite the protection flag.
      if (sheetProtected && effectiveFmt?.locked === false)
        paintLockMarker(ctx, bounds, theme, rtl);

      // Error / validation triangles. Error wins over validation when both
      // would apply (an error-kind value already implies the data is bad —
      // the green triangle conveys that without piling on a red one). The
      // ignoredErrors set suppresses both kinds for the cell once the user
      // dismisses the popover via the "Ignore" action.
      const cellKey = key;
      if (!ignored.has(cellKey)) {
        const validation = fmt?.validation ?? rangeValidationAt(format, cellAddr);
        if (detectErrorKind(value)) {
          const hit = paintErrorTriangle(ctx, bounds, ERROR_TRIANGLE_COLOR, rtl);
          triangleHits.push({ rect: hit, addr: cellAddr, kind: 'error' });
        } else if (validation && detectValidationViolation(value, validation, getResolver())) {
          const hit = paintValidationTriangle(ctx, bounds, VALIDATION_TRIANGLE_COLOR, rtl);
          triangleHits.push({ rect: hit, addr: cellAddr, kind: 'validation' });
        }
      }
    }
  }
  setErrorTriangleHits(triangleHits);
}
