// Live preview for the Format Cells dialog: projects the draft onto the
// shared preview cell plus the font, fill and border tabs' sample boxes.
// Read-only with respect to the draft.

import { formatNumber } from '../commands/format.js';
import type { CellValue } from '../engine/types.js';
import { formatCell, formatGeneralNumber } from '../engine/value.js';
import type { CellAlign, CellBorderSide, CellVAlign, FillPattern } from '../store/store.js';
import type { DraftState, NumberCategory, SideKey } from './format-dialog-model.js';
import { activeDraftSide, computeDialogNumFmt } from './format-dialog-state.js';

export interface FormatPreviewRefs {
  preview: HTMLDivElement;
  previewCell: HTMLDivElement;
  fontPreviewBox: HTMLDivElement;
  fillSample: HTMLDivElement;
  borderVisualPreview: HTMLDivElement;
  visualSideButtons: Map<SideKey, HTMLButtonElement[]>;
}

export interface FormatPreviewInput {
  readonly refs: FormatPreviewRefs;
  readonly draft: DraftState;
  getLocale(): string;
  defaultPatternFor(category: NumberCategory): string;
  /** Active cell value; null while editing a differential format, where a
   *  representative synthetic sample stands in. */
  previewValue(): CellValue | null;
}

export const fillPatternImage = (pattern: FillPattern | undefined, color = '#000000'): string => {
  switch (pattern) {
    case 'gray0625':
      return `radial-gradient(${color} 0.4px, transparent 0.4px)`;
    case 'gray125':
      return `radial-gradient(${color} 0.6px, transparent 0.6px)`;
    case 'gray25':
      return `radial-gradient(${color} 1px, transparent 1px)`;
    case 'gray50':
      return `repeating-linear-gradient(45deg, ${color} 0 2px, transparent 2px 4px)`;
    case 'gray75':
      return `repeating-linear-gradient(45deg, ${color} 0 3px, transparent 3px 4px)`;
    case 'horizontal':
    case 'darkHorizontal':
      return `repeating-linear-gradient(0deg, ${color} 0 1px, transparent 1px 4px)`;
    case 'lightHorizontal':
      return `repeating-linear-gradient(0deg, ${color} 0 1px, transparent 1px 7px)`;
    case 'vertical':
    case 'darkVertical':
      return `repeating-linear-gradient(90deg, ${color} 0 1px, transparent 1px 4px)`;
    case 'lightVertical':
      return `repeating-linear-gradient(90deg, ${color} 0 1px, transparent 1px 7px)`;
    case 'diagonalDown':
    case 'darkDown':
      return `repeating-linear-gradient(45deg, ${color} 0 1px, transparent 1px 5px)`;
    case 'diagonalUp':
    case 'darkUp':
      return `repeating-linear-gradient(135deg, ${color} 0 1px, transparent 1px 5px)`;
    case 'lightDown':
      return `repeating-linear-gradient(45deg, ${color} 0 1px, transparent 1px 9px)`;
    case 'lightUp':
      return `repeating-linear-gradient(135deg, ${color} 0 1px, transparent 1px 9px)`;
    case 'darkGrid':
    case 'lightGrid': {
      const step = pattern === 'darkGrid' ? 4 : 7;
      return `repeating-linear-gradient(0deg, ${color} 0 1px, transparent 1px ${step}px), repeating-linear-gradient(90deg, ${color} 0 1px, transparent 1px ${step}px)`;
    }
    case 'darkTrellis':
    case 'lightTrellis': {
      const step = pattern === 'darkTrellis' ? 6 : 9;
      return `repeating-linear-gradient(45deg, ${color} 0 1px, transparent 1px ${step}px), repeating-linear-gradient(135deg, ${color} 0 1px, transparent 1px ${step}px)`;
    }
    default:
      return '';
  }
};

const cssHorizontalAlign = (align: CellAlign | undefined): CSSStyleDeclaration['textAlign'] => {
  switch (align) {
    case 'center':
    case 'centerContinuous':
      return 'center';
    case 'right':
      return 'right';
    case 'justify':
    case 'distributed':
      return 'justify';
    default:
      return 'left';
  }
};

const cssVerticalJustify = (
  align: CellVAlign | undefined,
): CSSStyleDeclaration['justifyContent'] => {
  switch (align) {
    case 'top':
      return 'flex-start';
    case 'bottom':
      return 'flex-end';
    default:
      return 'center';
  }
};

const cssBorder = (s: CellBorderSide | undefined): string => {
  if (!s) return '0 solid transparent';
  const cfg = typeof s === 'object' ? s : { style: 'thin' as const };
  const widthPx = cfg.style === 'thick' ? 3 : cfg.style === 'medium' ? 2 : 1;
  const cssStyle =
    cfg.style === 'dashed'
      ? 'dashed'
      : cfg.style === 'dotted'
        ? 'dotted'
        : cfg.style === 'double'
          ? 'double'
          : 'solid';
  const cssColor = (typeof s === 'object' && s.color) || 'currentColor';
  const w = cfg.style === 'double' ? Math.max(widthPx, 3) : widthPx;
  return `${w}px ${cssStyle} ${cssColor}`;
};

/** Returns the preview renderer; call it after every draft change. */
export function createFormatPreview(input: FormatPreviewInput): () => void {
  const { refs, draft, getLocale, defaultPatternFor } = input;
  const { preview, previewCell, fontPreviewBox, fillSample, borderVisualPreview } = refs;
  return (): void => {
    const cssFontVertAlign =
      draft.fontVertAlign === 'superscript'
        ? 'super'
        : draft.fontVertAlign === 'subscript'
          ? 'sub'
          : '';
    const applyFontPreview = (el: HTMLElement): void => {
      el.style.fontWeight = draft.bold ? 'bold' : 'normal';
      el.style.fontStyle = draft.italic ? 'italic' : 'normal';
      const decos: string[] = [];
      if (draft.underline) decos.push('underline');
      if (draft.strike) decos.push('line-through');
      el.style.textDecoration = decos.length > 0 ? decos.join(' ') : 'none';
      el.style.textDecorationStyle =
        draft.underline === 'double' || draft.underline === 'doubleAccounting' ? 'double' : '';
      el.style.fontFamily = draft.fontFamily || '';
      el.style.fontSize = draft.fontSize !== undefined ? `${draft.fontSize}px` : '';
      el.style.verticalAlign = cssFontVertAlign;
      el.style.color = draft.color ?? '';
    };
    const applyFillPreview = (el: HTMLElement): void => {
      el.style.backgroundColor = draft.fill ?? '';
      el.style.backgroundImage = fillPatternImage(draft.fillPattern, draft.fillPatternColor);
      el.style.backgroundSize =
        draft.fillPattern === 'gray125' || draft.fillPattern === 'gray25' ? '4px 4px' : '';
    };

    preview.style.fontWeight = draft.bold ? 'bold' : 'normal';
    preview.style.fontStyle = draft.italic ? 'italic' : 'normal';
    const decos: string[] = [];
    if (draft.underline) decos.push('underline');
    if (draft.strike) decos.push('line-through');
    preview.style.textDecoration = decos.length > 0 ? decos.join(' ') : 'none';
    preview.style.textDecorationStyle =
      draft.underline === 'double' || draft.underline === 'doubleAccounting' ? 'double' : '';
    preview.style.verticalAlign = cssFontVertAlign;
    preview.style.textAlign = cssHorizontalAlign(draft.align);
    applyFontPreview(previewCell);
    applyFontPreview(fontPreviewBox);
    previewCell.style.textAlign = cssHorizontalAlign(draft.align);
    previewCell.style.direction = draft.textDirection === 'context' ? '' : draft.textDirection;
    applyFillPreview(previewCell);
    applyFillPreview(fillSample);
    previewCell.style.whiteSpace = draft.wrap ? 'pre-wrap' : 'nowrap';
    previewCell.style.fontSize = draft.shrinkToFit
      ? `${Math.max(8, Math.round((draft.fontSize ?? 13) * 0.85))}px`
      : draft.fontSize !== undefined
        ? `${draft.fontSize}px`
        : '';
    previewCell.style.justifyContent = cssVerticalJustify(draft.vAlign);
    previewCell.style.borderTop = cssBorder(draft.borders.top);
    previewCell.style.borderRight = cssBorder(draft.borders.right);
    previewCell.style.borderBottom = cssBorder(draft.borders.bottom);
    previewCell.style.borderLeft = cssBorder(draft.borders.left);
    borderVisualPreview.style.borderTop = cssBorder(draft.borders.top);
    borderVisualPreview.style.borderRight = cssBorder(draft.borders.right);
    borderVisualPreview.style.borderBottom = cssBorder(draft.borders.bottom);
    borderVisualPreview.style.borderLeft = cssBorder(draft.borders.left);
    borderVisualPreview.classList.toggle(
      'fc-fmtdlg__border-preview--diag-down',
      !!draft.borders.diagonalDown,
    );
    borderVisualPreview.classList.toggle(
      'fc-fmtdlg__border-preview--diag-up',
      !!draft.borders.diagonalUp,
    );
    const diagSide =
      draft.borders.diagonalDown || draft.borders.diagonalUp || activeDraftSide(draft);
    const diagCfg = typeof diagSide === 'object' ? diagSide : { style: 'thin' as const };
    const diagColor = (typeof diagSide === 'object' && diagSide.color) || 'currentColor';
    const diagWidth = diagCfg.style === 'thick' ? 3 : diagCfg.style === 'medium' ? 2 : 1;
    borderVisualPreview.style.setProperty('--fc-fmtdlg-border-diag-color', diagColor);
    borderVisualPreview.style.setProperty('--fc-fmtdlg-border-diag-width', `${diagWidth}px`);
    for (const [key, buttons] of refs.visualSideButtons) {
      for (const btn of buttons) {
        btn.setAttribute('aria-pressed', draft.borders[key] ? 'true' : 'false');
      }
    }

    const numFmt = computeDialogNumFmt(draft, defaultPatternFor);
    // Differential-format editing has no cell value to preview, so retain a
    // representative sample there. Normal Format Cells previews the active
    // cell, including text, booleans, errors, and blanks.
    const isDateLike =
      numFmt.kind === 'date' || numFmt.kind === 'time' || numFmt.kind === 'datetime';
    const syntheticSampleValue =
      draft.numberCategory === 'fraction'
        ? 1.25
        : (draft.numberCategory === 'fixed' || draft.numberCategory === 'currency') &&
            draft.negativeStyle !== 'minus'
          ? -1234
          : isDateLike || draft.numberCategory === 'currency' || draft.numberCategory === 'special'
            ? 10
            : 12345;
    const cellValue = input.previewValue();
    const synthetic = cellValue === null;
    const value = synthetic ? { kind: 'number' as const, value: syntheticSampleValue } : cellValue;
    const numericText =
      value.kind === 'number'
        ? synthetic || numFmt.kind !== 'general'
          ? formatNumber(value.value, numFmt, getLocale())
          : formatGeneralNumber(value.value, getLocale(), { useGrouping: false })
        : formatCell(value, getLocale());
    previewCell.textContent = numericText;
    if (!draft.color && value.kind === 'number' && value.value < 0) {
      previewCell.style.color =
        draft.negativeStyle === 'red' || draft.negativeStyle === 'red-parens' ? '#c00000' : '';
    }
  };
}
