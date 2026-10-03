import type { SelectionFormatPlan } from '../commands/format.js';
import { formatWithPending } from '../store/pending-format.js';
import type {
  CellBorderSide,
  CellBorders,
  CellFormat,
  CellValidation,
  NumFmt,
  State,
} from '../store/store.js';
import {
  type BorderStyleKey,
  type DraftState,
  defaultCurrencySymbolFor,
  FRACTION_PATTERNS,
  type NumberCategory,
  type SideKey,
} from './format-dialog-model.js';

export type FormatDialogField =
  | 'numFmt'
  | 'align'
  | 'vAlign'
  | 'wrap'
  | 'justifyLastLine'
  | 'shrinkToFit'
  | 'indent'
  | 'rotation'
  | 'textDirection'
  | 'bold'
  | 'italic'
  | 'underline'
  | 'strike'
  | 'fontVertAlign'
  | 'fontFamily'
  | 'fontSize'
  | 'color'
  | 'fill'
  | 'fillPattern'
  | 'fillPatternColor'
  | 'border.top'
  | 'border.right'
  | 'border.bottom'
  | 'border.left'
  | 'border.diagonalDown'
  | 'border.diagonalUp'
  | 'locked'
  | 'formulaHidden'
  | 'hyperlink'
  | 'comment'
  | 'validation';

export interface DialogFormatSummary {
  readonly activeFormat: CellFormat;
  readonly mixed: Set<FormatDialogField>;
}

const cloneFormat = (format: CellFormat | undefined): CellFormat =>
  format ? structuredClone(format) : {};

const normalizedBorderSide = (side: CellBorderSide | undefined): unknown => {
  if (!side) return false;
  if (side === true) return { style: 'thin' };
  return { style: side.style, color: side.color ?? undefined };
};

const normalizedFieldValue = (
  format: CellFormat | undefined,
  field: FormatDialogField,
): unknown => {
  const value = format ?? {};
  switch (field) {
    case 'numFmt':
      return value.numFmt ?? { kind: 'general' };
    case 'align':
    case 'vAlign':
      return value[field];
    case 'wrap':
    case 'justifyLastLine':
    case 'shrinkToFit':
    case 'bold':
    case 'italic':
    case 'strike':
    case 'formulaHidden':
      return value[field] === true;
    case 'indent':
    case 'rotation':
      return value[field] ?? 0;
    case 'textDirection':
      return value.textDirection ?? 'context';
    case 'underline':
      return value.underline === true ? 'single' : (value.underline ?? false);
    case 'fontVertAlign':
    case 'fontSize':
      return value[field];
    case 'fontFamily':
      return value.fontFamily ?? '';
    case 'color':
    case 'fill':
    case 'fillPattern':
    case 'fillPatternColor':
      return value[field] ?? undefined;
    case 'locked':
      return value.locked !== false;
    case 'hyperlink':
      return value.hyperlink ?? '';
    case 'comment':
      return value.comment ?? '';
    case 'validation':
      return value.validation;
    case 'border.top':
    case 'border.right':
    case 'border.bottom':
    case 'border.left':
    case 'border.diagonalDown':
    case 'border.diagonalUp':
      return normalizedBorderSide(value.borders?.[field.slice(7) as keyof CellBorders]);
  }
};

const sameNormalized = (left: unknown, right: unknown): boolean => {
  if (Object.is(left, right)) return true;
  if (left === null || right === null || typeof left !== 'object' || typeof right !== 'object') {
    return false;
  }
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right) || left.length !== right.length) return false;
    return left.every((value, index) => sameNormalized(value, right[index]));
  }
  const leftRecord = left as Record<string, unknown>;
  const rightRecord = right as Record<string, unknown>;
  const leftKeys = Object.keys(leftRecord);
  const rightKeys = Object.keys(rightRecord);
  if (leftKeys.length !== rightKeys.length) return false;
  return leftKeys.every(
    (key) => Object.hasOwn(rightRecord, key) && sameNormalized(leftRecord[key], rightRecord[key]),
  );
};

export function summarizeDialogFormats(
  state: State,
  plan: SelectionFormatPlan | null,
): DialogFormatSummary {
  const activeFormat = cloneFormat(formatWithPending(state, state.selection.active));
  const mixed = new Set<FormatDialogField>();
  if (!plan || plan.cells.length === 0) return { activeFormat, mixed };
  const fields: FormatDialogField[] = [
    'numFmt',
    'align',
    'vAlign',
    'wrap',
    'justifyLastLine',
    'shrinkToFit',
    'indent',
    'rotation',
    'textDirection',
    'bold',
    'italic',
    'underline',
    'strike',
    'fontVertAlign',
    'fontFamily',
    'fontSize',
    'color',
    'fill',
    'fillPattern',
    'fillPatternColor',
    'border.top',
    'border.right',
    'border.bottom',
    'border.left',
    'border.diagonalDown',
    'border.diagonalUp',
    'locked',
    'formulaHidden',
    'hyperlink',
    'comment',
    'validation',
  ];
  for (const field of fields) {
    const expected = normalizedFieldValue(activeFormat, field);
    if (
      plan.cells.some(
        (addr) =>
          !sameNormalized(normalizedFieldValue(formatWithPending(state, addr), field), expected),
      )
    ) {
      mixed.add(field);
    }
  }
  return { activeFormat, mixed };
}

const defaultPatternFor = (category: NumberCategory): string => {
  switch (category) {
    case 'date':
      return 'yyyy-mm-dd';
    case 'time':
      return 'HH:MM:SS';
    case 'fraction':
      return '# ?/?';
    case 'special':
      return '000';
    case 'custom':
      return '0.00';
    default:
      return '';
  }
};

export function buildTouchedDialogPatch(
  draft: DraftState,
  touched: ReadonlySet<FormatDialogField>,
  defaultPattern: (category: NumberCategory) => string = defaultPatternFor,
): Partial<CellFormat> {
  const patch: Partial<CellFormat> = {};
  const set = <K extends keyof CellFormat>(key: K, value: CellFormat[K]): void => {
    if (touched.has(key as FormatDialogField)) patch[key] = value;
  };
  set('numFmt', computeDialogNumFmt(draft, defaultPattern));
  set('align', draft.align);
  set('vAlign', draft.vAlign);
  set('wrap', draft.wrap);
  set('justifyLastLine', draft.justifyLastLine);
  set('shrinkToFit', draft.shrinkToFit);
  set('indent', draft.indent > 0 ? draft.indent : undefined);
  set('rotation', draft.rotation !== 0 ? draft.rotation : undefined);
  set('textDirection', draft.textDirection === 'context' ? undefined : draft.textDirection);
  set('bold', draft.bold);
  set('italic', draft.italic);
  set('underline', draft.underline);
  set('strike', draft.strike);
  set('fontVertAlign', draft.fontVertAlign);
  set('fontFamily', draft.fontFamily || undefined);
  set('fontSize', draft.fontSize);
  set('color', draft.color);
  set('fill', draft.fill);
  set('fillPattern', draft.fillPattern);
  set('fillPatternColor', draft.fillPatternColor);
  set('locked', draft.locked);
  set('formulaHidden', draft.formulaHidden);
  if (touched.has('hyperlink')) patch.hyperlink = draft.hyperlink.trim() || undefined;
  if (touched.has('comment')) {
    patch.comment = draft.comment || undefined;
    if (!draft.comment) patch.commentAuthor = undefined;
  }
  if (touched.has('validation')) {
    const lines = draft.validationList
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter((line) => line.length > 0);
    patch.validation = computeDialogValidation(draft, lines);
  }
  const border: CellBorders = {};
  const sides: SideKey[] = ['top', 'right', 'bottom', 'left', 'diagonalDown', 'diagonalUp'];
  for (const side of sides) {
    if (touched.has(`border.${side}` as FormatDialogField)) {
      border[side] = draft.borders[side] ?? false;
    }
  }
  if (Object.keys(border).length > 0) patch.borders = border;
  return patch;
}

export function makeEmptyDraft(formatLocale: string): DraftState {
  return {
    numFmt: undefined,
    numberCategory: 'general',
    decimals: 2,
    thousands: false,
    negativeStyle: 'minus',
    currencySymbol: defaultCurrencySymbolFor(formatLocale),
    pattern: '',
    align: undefined,
    vAlign: undefined,
    wrap: false,
    justifyLastLine: false,
    shrinkToFit: false,
    indent: 0,
    rotation: 0,
    textDirection: 'context',
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    fontVertAlign: undefined,
    fontFamily: '',
    fontSize: undefined,
    color: undefined,
    fill: undefined,
    fillPattern: undefined,
    fillPatternColor: undefined,
    borders: {},
    borderStyle: 'thin',
    borderColor: undefined,
    hyperlink: '',
    originalHyperlink: '',
    hyperlinkDisplay: undefined,
    hyperlinkTooltip: undefined,
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
    validationShowInputMessage: true,
    validationPromptTitle: '',
    validationPromptMessage: '',
    validationShowErrorMessage: true,
    validationErrorTitle: '',
    validationErrorMessage: '',
    validationShowDropdown: true,
    locked: true,
    formulaHidden: false,
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
        draft.thousands = fmt.numFmt.thousands === true;
        draft.negativeStyle = fmt.numFmt.negativeStyle ?? 'minus';
        break;
      case 'currency':
        draft.numberCategory = 'currency';
        draft.decimals = fmt.numFmt.decimals;
        draft.currencySymbol = fmt.numFmt.symbol ?? '$';
        draft.negativeStyle = fmt.numFmt.negativeStyle ?? 'minus';
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
        draft.numberCategory = 'custom';
        draft.pattern = fmt.numFmt.pattern;
        break;
      case 'special':
        draft.numberCategory = 'special';
        draft.pattern = fmt.numFmt.pattern;
        break;
      case 'text':
        draft.numberCategory = 'text';
        break;
      case 'custom':
        draft.numberCategory = FRACTION_PATTERNS.includes(
          fmt.numFmt.pattern as (typeof FRACTION_PATTERNS)[number],
        )
          ? 'fraction'
          : 'custom';
        draft.pattern = fmt.numFmt.pattern;
        break;
      default:
        draft.numberCategory = 'general';
    }
  } else {
    draft.numFmt = { kind: 'general' };
    draft.numberCategory = 'general';
    draft.decimals = 2;
    draft.thousands = false;
    draft.negativeStyle = 'minus';
    draft.currencySymbol = defaultCurrencySymbolFor(formatLocale);
    draft.pattern = '';
  }

  draft.align = fmt.align;
  draft.vAlign = fmt.vAlign;
  draft.wrap = !!fmt.wrap;
  draft.justifyLastLine = !!fmt.justifyLastLine;
  draft.shrinkToFit = !!fmt.shrinkToFit;
  draft.indent = fmt.indent ?? 0;
  draft.rotation = fmt.rotation ?? 0;
  draft.textDirection = fmt.textDirection ?? 'context';
  draft.bold = !!fmt.bold;
  draft.italic = !!fmt.italic;
  draft.underline = fmt.underline === true ? 'single' : (fmt.underline ?? false);
  draft.strike = !!fmt.strike;
  draft.fontVertAlign = fmt.fontVertAlign;
  draft.fontFamily = fmt.fontFamily ?? '';
  draft.fontSize = fmt.fontSize;
  draft.color = fmt.color;
  draft.fill = fmt.fill;
  draft.fillPattern = fmt.fillPattern;
  draft.fillPatternColor = fmt.fillPatternColor;
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
  draft.originalHyperlink = draft.hyperlink;
  draft.hyperlinkDisplay = fmt.hyperlinkDisplay;
  draft.hyperlinkTooltip = fmt.hyperlinkTooltip;
  draft.comment = fmt.comment ?? '';
  hydrateValidationDraft(draft, fmt.validation);
  draft.locked = fmt.locked !== false;
  draft.formulaHidden = fmt.formulaHidden === true;
}

export function activeDraftSide(draft: DraftState): CellBorderSide {
  return {
    style: draft.borderStyle,
    ...(draft.borderColor ? { color: draft.borderColor } : {}),
  };
}

export function setDraftSide(draft: DraftState, key: SideKey, on: boolean): CellBorders {
  const next: CellBorders = { ...draft.borders };
  if (key === 'diagonalDown' || key === 'diagonalUp') {
    if (!on) {
      next[key] = false;
      return next;
    }
    const otherKey = key === 'diagonalDown' ? 'diagonalUp' : 'diagonalDown';
    const shared = next[otherKey] || activeDraftSide(draft);
    next.diagonalDown = shared;
    next.diagonalUp = shared;
    return next;
  }
  if (on) next[key] = activeDraftSide(draft);
  else next[key] = false;
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
      return {
        kind: 'fixed',
        decimals: draft.decimals,
        ...(draft.thousands ? { thousands: true } : {}),
        ...(draft.negativeStyle !== 'minus' ? { negativeStyle: draft.negativeStyle } : {}),
      };
    case 'currency':
      return {
        kind: 'currency',
        decimals: draft.decimals,
        symbol: draft.currencySymbol,
        ...(draft.negativeStyle !== 'minus' ? { negativeStyle: draft.negativeStyle } : {}),
      };
    case 'percent':
      return { kind: 'percent', decimals: draft.decimals };
    case 'scientific':
      return { kind: 'scientific', decimals: draft.decimals };
    case 'accounting':
      return { kind: 'accounting', decimals: draft.decimals, symbol: draft.currencySymbol };
    case 'text':
      return { kind: 'text' };
    case 'special':
      return { kind: 'special', pattern: draft.pattern || defaultPatternFor('special') };
    case 'date':
      return { kind: 'date', pattern: draft.pattern || defaultPatternFor('date') };
    case 'time':
      return { kind: 'time', pattern: draft.pattern || defaultPatternFor('time') };
    case 'fraction':
      return { kind: 'custom', pattern: draft.pattern || defaultPatternFor('fraction') };
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
    ...(draft.validationShowInputMessage ? {} : { showInputMessage: false }),
    ...(draft.validationPromptTitle.trim()
      ? { promptTitle: draft.validationPromptTitle.trim() }
      : {}),
    ...(draft.validationPromptMessage.trim()
      ? { promptMessage: draft.validationPromptMessage.trim() }
      : {}),
    ...(draft.validationShowErrorMessage ? {} : { showErrorMessage: false }),
    ...(draft.validationErrorTitle.trim() ? { errorTitle: draft.validationErrorTitle.trim() } : {}),
    ...(draft.validationErrorMessage.trim()
      ? { errorMessage: draft.validationErrorMessage.trim() }
      : {}),
    ...(draft.validationShowDropdown ? {} : { showDropdown: false }),
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
    draft.validationShowInputMessage = true;
    draft.validationPromptTitle = '';
    draft.validationPromptMessage = '';
    draft.validationShowErrorMessage = true;
    draft.validationErrorTitle = '';
    draft.validationErrorMessage = '';
    draft.validationShowDropdown = true;
    return;
  }

  draft.validationKind = v.kind;
  draft.validationAllowBlank = v.allowBlank !== false;
  draft.validationErrorStyle = v.errorStyle ?? 'stop';
  draft.validationShowInputMessage = v.showInputMessage !== false;
  draft.validationPromptTitle = v.promptTitle ?? '';
  draft.validationPromptMessage = v.promptMessage ?? '';
  draft.validationShowErrorMessage = v.showErrorMessage !== false;
  draft.validationErrorTitle = v.errorTitle ?? '';
  draft.validationErrorMessage = v.errorMessage ?? '';
  draft.validationShowDropdown = v.showDropdown !== false;
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
  if (typeof s === 'object') return s.style;
  return 'thin';
}

function sideColor(s: CellBorderSide | undefined): string | undefined {
  if (!s) return undefined;
  if (typeof s === 'object') return s.color;
  return undefined;
}
