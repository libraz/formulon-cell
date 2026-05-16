import type { CellValue } from '../../engine/types.js';
import type { CellFormat } from '../../store/store.js';
import type { ResolvedTheme } from '../../theme/resolve.js';
import type { Rect } from '../geometry.js';

export interface CellPaintCtx {
  ctx: CanvasRenderingContext2D;
  theme: ResolvedTheme;
  bounds: Rect;
  value: CellValue;
  formula: string | null;
  isActive: boolean;
  isInRange: boolean;
  format?: CellFormat;
  /** When true and `formula` is non-null, paint the formula text instead of
   *  the evaluated value (the desktop-spreadsheet "Show Formulas" mode). */
  showFormulas?: boolean;
  /** Override the displayed string. Set by `paintCells` after consulting
   *  the cell registry (`inst.cells.registerFormatter`). When non-null
   *  the formatter wins over numFmt + default `formatCell`. Empty string
   *  is honored — render-blank-cell-still-padded scenarios. */
  displayOverride?: string | null;
  /** BCP 47 locale used for number/date formatting. */
  locale?: string;
}

export type TextVAlign = 'top' | 'middle' | 'bottom';

export interface TextMetricsBox {
  ascent: number;
  descent: number;
}
