import type {
  CellBorderSide,
  CellBorders,
  CellFormat,
  CellValidation,
  NumFmt,
} from '../store/store.js';
import {
  type BorderStyleKey,
  type DraftState,
  defaultCurrencySymbolFor,
  type NumberCategory,
  type SideKey,
} from './format-dialog-model.js';

export function makeEmptyDraft(formatLocale: string): DraftState {
  return {
    numFmt: undefined,
    numberCategory: 'general',
    decimals: 2,
    currencySymbol: defaultCurrencySymbolFor(formatLocale),
    pattern: '',
    align: undefined,
    vAlign: undefined,
    wrap: false,
    indent: 0,
    rotation: 0,
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    fontFamily: '',
    fontSize: undefined,
    color: undefined,
    fill: undefined,
    borders: {},
    borderStyle: 'thin',
    borderColor: undefined,
    hyperlink: '',
    comment: '',
    validationList: '',
    validationListSourceKind: 'literal',
    validationListRange: '',
    validationKind: 'none',
    validationOp: 'between',
    validationA: 0,
    validationB: 0,
    validationFormula: '',
    validationAllowBlank: true,
    validationErrorStyle: 'stop',
    locked: true,
  };
}

export function hydrateDraftFromFormat(
  draft: DraftState,
  fmt: CellFormat,
  formatLocale: string,
): void {
  if (fmt.numFmt) {
    draft.numFmt = fmt.numFmt;
    switch (fmt.numFmt.kind) {
      case 'fixed':
        draft.numberCategory = 'fixed';
        draft.decimals = fmt.numFmt.decimals;
        break;
      case 'currency':
        draft.numberCategory = 'currency';
        draft.decimals = fmt.numFmt.decimals;
        draft.currencySymbol = fmt.numFmt.symbol ?? '$';
        break;
      case 'percent':
        draft.numberCategory = 'percent';
        draft.decimals = fmt.numFmt.decimals;
        break;
      case 'scientific':
        draft.numberCategory = 'scientific';
        draft.decimals = fmt.numFmt.decimals;
        break;
      case 'accounting':
        draft.numberCategory = 'accounting';
        draft.decimals = fmt.numFmt.decimals;
        draft.currencySymbol = fmt.numFmt.symbol ?? '$';
        break;
      case 'date':
        draft.numberCategory = 'date';
        draft.pattern = fmt.numFmt.pattern;
        break;
      case 'time':
        draft.numberCategory = 'time';
        draft.pattern = fmt.numFmt.pattern;
        break;
      case 'datetime':
        draft.numberCategory = 'datetime';
        draft.pattern = fmt.numFmt.pattern;
        break;
      case 'text':
        draft.numberCategory = 'text';
        break;
      case 'custom':
        draft.numberCategory = 'custom';
        draft.pattern = fmt.numFmt.pattern;
        break;
      default:
        draft.numberCategory = 'general';
    }
  } else {
    draft.numFmt = { kind: 'general' };
    draft.numberCategory = 'general';
    draft.decimals = 2;
    draft.currencySymbol = defaultCurrencySymbolFor(formatLocale);
    draft.pattern = '';
  }

  draft.align = fmt.align;
  draft.vAlign = fmt.vAlign;
  draft.wrap = !!fmt.wrap;
  draft.indent = fmt.indent ?? 0;
  draft.rotation = fmt.rotation ?? 0;
  draft.bold = !!fmt.bold;
  draft.italic = !!fmt.italic;
  draft.underline = !!fmt.underline;
  draft.strike = !!fmt.strike;
  draft.fontFamily = fmt.fontFamily ?? '';
  draft.fontSize = fmt.fontSize;
  draft.color = fmt.color;
  draft.fill = fmt.fill;
  draft.borders = { ...(fmt.borders ?? {}) };

  const sides: SideKey[] = ['top', 'right', 'bottom', 'left', 'diagonalDown', 'diagonalUp'];
  let inheritedStyle: BorderStyleKey | null = null;
  let inheritedColor: string | undefined;
  for (const k of sides) {
    const s = draft.borders[k];
    const ss = sideStyle(s);
    if (ss && !inheritedStyle) inheritedStyle = ss;
    const cc = sideColor(s);
    if (cc && !inheritedColor) inheritedColor = cc;
  }
  draft.borderStyle = inheritedStyle ?? 'thin';
  draft.borderColor = inheritedColor;

  draft.hyperlink = fmt.hyperlink ?? '';
  draft.comment = fmt.comment ?? '';
  hydrateValidationDraft(draft, fmt.validation);
  draft.locked = fmt.locked !== false;
}

export function activeDraftSide(draft: DraftState): CellBorderSide {
  return {
    style: draft.borderStyle,
    ...(draft.borderColor ? { color: draft.borderColor } : {}),
  };
}

export function setDraftSide(draft: DraftState, key: SideKey, on: boolean): CellBorders {
  const next: CellBorders = { ...draft.borders };
  if (on) next[key] = activeDraftSide(draft);
  else next[key] = false;
  return next;
}

export function restyleDraftBorders(draft: DraftState): CellBorders {
  const next: CellBorders = {};
  const sides: SideKey[] = ['top', 'right', 'bottom', 'left', 'diagonalDown', 'diagonalUp'];
  for (const k of sides) {
    if (draft.borders[k]) next[k] = activeDraftSide(draft);
  }
  return next;
}

export function explicitDraftBorders(draft: DraftState): CellBorders {
  return {
    top: draft.borders.top ?? false,
    right: draft.borders.right ?? false,
    bottom: draft.borders.bottom ?? false,
    left: draft.borders.left ?? false,
    diagonalDown: draft.borders.diagonalDown ?? false,
    diagonalUp: draft.borders.diagonalUp ?? false,
  };
}

export function computeDialogNumFmt(
  draft: DraftState,
  defaultPatternFor: (cat: NumberCategory) => string,
): NumFmt {
  const cat = draft.numberCategory;
  switch (cat) {
    case 'general':
      return { kind: 'general' };
    case 'fixed':
      return { kind: 'fixed', decimals: draft.decimals };
    case 'currency':
      return { kind: 'currency', decimals: draft.decimals, symbol: draft.currencySymbol };
    case 'percent':
      return { kind: 'percent', decimals: draft.decimals };
    case 'scientific':
      return { kind: 'scientific', decimals: draft.decimals };
    case 'accounting':
      return { kind: 'accounting', decimals: draft.decimals, symbol: draft.currencySymbol };
    case 'text':
      return { kind: 'text' };
    case 'date':
      return { kind: 'date', pattern: draft.pattern || defaultPatternFor('date') };
    case 'time':
      return { kind: 'time', pattern: draft.pattern || defaultPatternFor('time') };
    case 'datetime':
      return { kind: 'datetime', pattern: draft.pattern || defaultPatternFor('datetime') };
    case 'custom':
      return { kind: 'custom', pattern: draft.pattern || defaultPatternFor('custom') };
  }
}

export function computeDialogValidation(
  draft: DraftState,
  lines: string[],
): CellValidation | undefined {
  const k = draft.validationKind;
  if (k === 'none') return undefined;
  const meta = {
    ...(draft.validationAllowBlank ? {} : { allowBlank: false }),
    ...(draft.validationErrorStyle !== 'stop' ? { errorStyle: draft.validationErrorStyle } : {}),
  };
  switch (k) {
    case 'list':
      if (draft.validationListSourceKind === 'range') {
        const ref = draft.validationListRange.trim().replace(/^=/, '');
        if (!ref) return undefined;
        return { kind: 'list', source: { ref }, ...meta };
      }
      if (lines.length === 0) return undefined;
      return { kind: 'list', source: lines, ...meta };
    case 'custom': {
      const formula = draft.validationFormula.trim();
      if (!formula) return undefined;
      return { kind: 'custom', formula, ...meta };
    }
    case 'whole':
    case 'decimal':
    case 'date':
    case 'time':
    case 'textLength': {
      const op = draft.validationOp;
      const a = draft.validationA;
      if (op === 'between' || op === 'notBetween') {
        return { kind: k, op, a, b: draft.validationB, ...meta };
      }
      return { kind: k, op, a, ...meta };
    }
  }
}

function hydrateValidationDraft(draft: DraftState, validation: CellValidation | undefined): void {
  const v = validation;
  if (!v) {
    draft.validationKind = 'none';
    draft.validationList = '';
    draft.validationListRange = '';
    draft.validationListSourceKind = 'literal';
    draft.validationFormula = '';
    draft.validationOp = 'between';
    draft.validationA = 0;
    draft.validationB = 0;
    draft.validationAllowBlank = true;
    draft.validationErrorStyle = 'stop';
    return;
  }

  draft.validationKind = v.kind;
  draft.validationAllowBlank = v.allowBlank !== false;
  draft.validationErrorStyle = v.errorStyle ?? 'stop';
  if (v.kind === 'list') {
    if (Array.isArray(v.source)) {
      draft.validationListSourceKind = 'literal';
      draft.validationList = v.source.join('\n');
      draft.validationListRange = '';
    } else {
      draft.validationListSourceKind = 'range';
      draft.validationList = '';
      draft.validationListRange = v.source.ref;
    }
  } else {
    draft.validationList = '';
    draft.validationListRange = '';
    draft.validationListSourceKind = 'literal';
  }
  draft.validationFormula = v.kind === 'custom' ? v.formula : '';
  if (
    v.kind === 'whole' ||
    v.kind === 'decimal' ||
    v.kind === 'date' ||
    v.kind === 'time' ||
    v.kind === 'textLength'
  ) {
    draft.validationOp = v.op;
    draft.validationA = v.a;
    draft.validationB = v.b ?? v.a;
  } else {
    draft.validationOp = 'between';
    draft.validationA = 0;
    draft.validationB = 0;
  }
}

function sideStyle(s: CellBorderSide | undefined): BorderStyleKey | null {
  if (!s) return null;
  if (typeof s === 'object') {
    switch (s.style) {
      case 'thin':
      case 'medium':
      case 'thick':
      case 'dashed':
      case 'dotted':
      case 'double':
        return s.style;
      case 'hair':
        return 'thin';
      case 'mediumDashed':
      case 'dashDot':
      case 'mediumDashDot':
      case 'dashDotDot':
      case 'mediumDashDotDot':
      case 'slantDashDot':
        return 'dashed';
      default:
        return 'thin';
    }
  }
  return 'thin';
}

function sideColor(s: CellBorderSide | undefined): string | undefined {
  if (!s) return undefined;
  if (typeof s === 'object') return s.color;
  return undefined;
}
