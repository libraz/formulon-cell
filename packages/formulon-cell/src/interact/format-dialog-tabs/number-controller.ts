// Number tab behaviour for the Format Cells dialog: category list, decimals,
// separators, currency symbol, negative-number styles and the pattern
// presets/listbox. The tab's DOM is built inline by `format-dialog-view.ts`.

import { formatNumber } from '../../commands/format.js';
import type { Strings } from '../../i18n/strings.js';
import type { NegativeStyle } from '../../store/store.js';
import { appendDialogSelectOptions } from '../../toolbar/dialogs/form-controls.js';
import { appendDialogOptionButton } from '../dialog-shell.js';
import {
  defaultCurrencySymbolFor,
  defaultPatternForLocale,
  type NumberCategory,
  normalizeFormatLocale,
  patternPresetsFor,
} from '../format-dialog-model.js';
import { computeDialogNumFmt } from '../format-dialog-state.js';
import type { createFormatDialogView } from '../format-dialog-view.js';
import {
  type FormatTabContext,
  type FormatTabController,
  markMixed,
  markMixedCheck,
} from './controller.js';

export type NumberTabRefs = Pick<
  ReturnType<typeof createFormatDialogView>,
  | 'tabPanels'
  | 'catList'
  | 'catDefs'
  | 'catButtons'
  | 'decimalsRow'
  | 'decimalsInput'
  | 'thousandsCk'
  | 'symbolRow'
  | 'symbolSelect'
  | 'patternPresetRow'
  | 'patternPresetSelect'
  | 'patternListWrap'
  | 'patternList'
  | 'patternRow'
  | 'patternInput'
  | 'localeRow'
  | 'localeSelect'
  | 'calendarRow'
  | 'negativeList'
  | 'negativeOptions'
  | 'numberSummaryTitle'
  | 'numberSummaryDesc'
>;

export function numberCategoryDescription(cat: NumberCategory, t: Strings['formatDialog']): string {
  switch (cat) {
    case 'fixed':
      return t.descFixed;
    case 'currency':
      return t.descCurrency;
    case 'accounting':
      return t.descAccounting;
    case 'percent':
      return t.descPercent;
    case 'scientific':
      return t.descScientific;
    case 'date':
      return t.descDate;
    case 'time':
      return t.descTime;
    case 'fraction':
      return t.descFraction;
    case 'text':
      return t.descText;
    case 'special':
      return t.descOther;
    case 'custom':
      return t.descCustom;
    default:
      return t.descGeneral;
  }
}

// Pattern preview values per category — pick a value that exercises the
// formatting rules so users see day-of-week, AM/PM, etc.
const patternSampleValue = (cat: NumberCategory): number => {
  switch (cat) {
    case 'date':
      return 41348.5625; // 2013-03-14 13:30
    case 'fraction':
      return 1.25;
    case 'time':
      return 0.5625; // 13:30:00
    case 'special':
      return 12345;
    default:
      return 12345;
  }
};

/** `syncHintBar` refreshes the dialog-wide hint bar, which shows the active
 *  category's description while this tab is selected. */
export function attachNumberTab(
  ctx: FormatTabContext,
  refs: NumberTabRefs,
  syncHintBar: () => void,
): FormatTabController {
  const { draft, t, on, touch, isMixed } = ctx;
  const {
    tabPanels,
    catList,
    catDefs,
    catButtons,
    decimalsRow,
    decimalsInput,
    thousandsCk,
    symbolRow,
    symbolSelect,
    patternPresetRow,
    patternPresetSelect,
    patternListWrap,
    patternList,
    patternRow,
    patternInput,
    localeRow,
    localeSelect,
    calendarRow,
    negativeList,
    negativeOptions,
    numberSummaryTitle,
    numberSummaryDesc,
  } = refs;
  const defaultPatternFor = (cat: NumberCategory): string =>
    defaultPatternForLocale(cat, ctx.getLocale());

  const syncPatternListItems = (
    patterns: string[],
    current: string,
    specialLabels: string[],
  ): void => {
    const cat = draft.numberCategory;
    const isListbox = cat === 'date' || cat === 'time' || cat === 'fraction' || cat === 'special';
    if (!isListbox) {
      patternList.replaceChildren();
      return;
    }
    const sample = patternSampleValue(cat);
    const locale = ctx.getLocale();
    patternList.replaceChildren();
    for (const [index, pattern] of patterns.entries()) {
      let label = '';
      if (cat === 'special') {
        label = specialLabels[index] ?? pattern;
      } else {
        try {
          label = formatNumber(
            sample,
            cat === 'date'
              ? { kind: 'date', pattern }
              : cat === 'time'
                ? { kind: 'time', pattern }
                : { kind: 'custom', pattern },
            locale,
          );
        } catch {
          label = pattern;
        }
        label = label || pattern;
      }
      appendDialogOptionButton(patternList, {
        label,
        baseClass: 'fc-fmtdlg__pattern-item',
        datasetKey: 'fcPattern',
        value: pattern,
        selected: pattern === current,
      });
    }
  };

  const syncPatternPresetOptions = (): void => {
    const cat = draft.numberCategory;
    const specialLabels = t.specialFormatLabels.split('\n');
    const patterns =
      cat === 'date' ||
      cat === 'time' ||
      cat === 'fraction' ||
      cat === 'special' ||
      cat === 'custom'
        ? [...patternPresetsFor(ctx.getLocale())[cat]]
        : [];
    const current = draft.pattern || defaultPatternFor(cat);
    if (current && !patterns.includes(current)) patterns.unshift(current);
    patternPresetSelect.replaceChildren();
    appendDialogSelectOptions(
      patternPresetSelect,
      patterns.map((pattern, index) => ({
        value: pattern,
        label: cat === 'special' ? (specialLabels[index] ?? pattern) : pattern,
      })),
    );
    patternPresetSelect.value = current;
    syncPatternListItems(patterns, current, specialLabels);
  };

  const syncNegativeSamples = (): void => {
    const cat = draft.numberCategory;
    const items = negativeOptions.querySelectorAll<HTMLButtonElement>('[data-fc-negative-style]');
    for (const item of items) {
      const style = item.dataset.fcNegativeStyle as NegativeStyle | undefined;
      if (!style) continue;
      const sampleFmt = computeDialogNumFmt(
        { ...draft, numberCategory: cat, negativeStyle: style },
        defaultPatternFor,
      );
      item.textContent = formatNumber(-1234, sampleFmt, ctx.getLocale());
    }
  };

  const syncNumberControlsVisibility = (): void => {
    const cat = draft.numberCategory;
    tabPanels.get('number')?.setAttribute('data-number-category', cat);
    const decimalsCats = new Set<NumberCategory>([
      'fixed',
      'currency',
      'percent',
      'scientific',
      'accounting',
    ]);
    const symbolCats = new Set<NumberCategory>(['currency', 'accounting']);
    const listboxCats = new Set<NumberCategory>(['date', 'time', 'fraction', 'special']);
    decimalsRow.hidden = !decimalsCats.has(cat);
    thousandsCk.wrap.hidden = cat !== 'fixed';
    symbolRow.hidden = !symbolCats.has(cat);
    // For date/time-like categories use the Office365-style clickable
    // listbox; only the custom category keeps the dropdown of code-style
    // presets.
    patternPresetRow.hidden = cat !== 'custom';
    patternListWrap.hidden = !listboxCats.has(cat);
    patternRow.hidden = cat !== 'custom';
    localeRow.hidden = cat !== 'date' && cat !== 'time' && cat !== 'special';
    localeSelect.value = normalizeFormatLocale(ctx.getLocale()).startsWith('ja') ? 'ja' : 'en';
    calendarRow.hidden = cat !== 'date';
    negativeList.hidden = cat !== 'fixed' && cat !== 'currency';
    const active = catDefs.find((c) => c.id === cat);
    numberSummaryTitle.textContent = active?.label ?? t.catGeneral;
    // Description moved to the hint bar; keep the in-controls slot empty so
    // it does not push the layout.
    numberSummaryDesc.textContent = '';
    syncNegativeSamples();
    syncHintBar();
  };

  const sync = (): void => {
    for (const [id, btn] of catButtons) {
      btn.setAttribute('aria-selected', id === draft.numberCategory ? 'true' : 'false');
      btn.tabIndex = id === draft.numberCategory ? 0 : -1;
    }
    decimalsInput.value = String(draft.decimals);
    thousandsCk.input.checked = draft.thousands;
    for (const item of negativeOptions.querySelectorAll<HTMLButtonElement>(
      '[data-fc-negative-style]',
    )) {
      item.setAttribute(
        'aria-selected',
        item.dataset.fcNegativeStyle === draft.negativeStyle ? 'true' : 'false',
      );
    }
    symbolSelect.value = draft.currencySymbol;
    patternInput.value = draft.pattern;
    if (!draft.pattern) {
      patternInput.placeholder = defaultPatternFor(draft.numberCategory) || t.patternPlaceholder;
    } else {
      patternInput.placeholder = t.patternPlaceholder;
    }
    syncPatternPresetOptions();
    syncNumberControlsVisibility();
  };

  const syncMixed = (): void => {
    markMixedCheck(thousandsCk.input, isMixed('numFmt'), 'numFmt');
    if (!isMixed('numFmt')) return;
    for (const button of catButtons.values()) button.setAttribute('aria-selected', 'false');
    markMixed(catList, true, 'numFmt');
    decimalsInput.value = '';
    markMixed(decimalsInput, true, 'numFmt');
    symbolSelect.value = '';
    symbolSelect.selectedIndex = -1;
    markMixed(symbolSelect, true, 'numFmt');
    for (const button of negativeOptions.querySelectorAll<HTMLButtonElement>(
      '[data-fc-negative-style]',
    )) {
      button.setAttribute('aria-selected', 'false');
    }
    markMixed(negativeOptions, true, 'numFmt');
    patternInput.value = '';
    patternPresetSelect.value = '';
    for (const button of patternList.querySelectorAll<HTMLButtonElement>('[data-fc-pattern]')) {
      button.setAttribute('aria-selected', 'false');
    }
    markMixed(patternList, true, 'numFmt');
    markMixed(patternInput, true, 'numFmt');
    markMixed(patternPresetSelect, true, 'numFmt');
  };

  // ── Events ─────────────────────────────────────────────────────────────
  const setNumberCategory = (id: NumberCategory): void => {
    touch('numFmt');
    const previous = draft.numberCategory;
    draft.numberCategory = id;
    if (previous !== id) {
      const fallback = defaultPatternFor(id);
      if (fallback) draft.pattern = fallback;
      if (
        (id === 'currency' || id === 'accounting') &&
        previous !== 'currency' &&
        previous !== 'accounting'
      ) {
        draft.currencySymbol = defaultCurrencySymbolFor(ctx.getLocale());
      }
    }
    ctx.syncControls();
    ctx.renderPreview();
  };

  const onCatClick = (e: MouseEvent): void => {
    const target = e.target as HTMLElement;
    const btn = target.closest('button[data-fc-cat]') as HTMLButtonElement | null;
    if (!btn) return;
    const id = btn.dataset.fcCat as NumberCategory | undefined;
    if (!id) return;
    setNumberCategory(id);
    btn.focus();
  };

  const focusCategoryByIndex = (idx: number): void => {
    const categories = Array.from(catButtons.keys());
    const next = categories[(idx + categories.length) % categories.length];
    if (!next) return;
    setNumberCategory(next);
    catButtons.get(next)?.focus();
  };

  const onCatKeyDown = (e: KeyboardEvent): void => {
    const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('button[data-fc-cat]');
    if (!btn) return;
    const id = btn.dataset.fcCat as NumberCategory | undefined;
    const categories = Array.from(catButtons.keys());
    const idx = id ? categories.indexOf(id) : -1;
    if (idx < 0) return;
    if (e.key === 'ArrowDown' || e.key === 'ArrowRight') {
      e.preventDefault();
      focusCategoryByIndex(idx + 1);
    } else if (e.key === 'ArrowUp' || e.key === 'ArrowLeft') {
      e.preventDefault();
      focusCategoryByIndex(idx - 1);
    } else if (e.key === 'Home') {
      e.preventDefault();
      focusCategoryByIndex(0);
    } else if (e.key === 'End') {
      e.preventDefault();
      focusCategoryByIndex(categories.length - 1);
    }
  };

  const onDecimalsInput = (): void => {
    touch('numFmt');
    const n = Number.parseInt(decimalsInput.value, 10);
    if (Number.isFinite(n)) draft.decimals = Math.max(0, Math.min(10, n));
    ctx.syncControls();
    syncNegativeSamples();
    ctx.renderPreview();
  };

  const onThousandsChange = (): void => {
    touch('numFmt');
    draft.thousands = thousandsCk.input.checked;
    ctx.syncControls();
    syncNegativeSamples();
    ctx.renderPreview();
  };

  const onNegativeStyleClick = (e: Event): void => {
    const item = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-fc-negative-style]');
    const style = item?.dataset.fcNegativeStyle as NegativeStyle | undefined;
    if (!style) return;
    touch('numFmt');
    draft.negativeStyle = style;
    ctx.syncControls();
    ctx.renderPreview();
  };

  const onSymbolChange = (): void => {
    touch('numFmt');
    draft.currencySymbol = symbolSelect.value;
    ctx.syncControls();
    syncNegativeSamples();
    ctx.renderPreview();
  };

  const onPatternInput = (): void => {
    touch('numFmt');
    draft.pattern = patternInput.value;
    ctx.syncControls();
    syncPatternPresetOptions();
    ctx.renderPreview();
  };

  const onPatternPresetChange = (): void => {
    touch('numFmt');
    draft.pattern = patternPresetSelect.value;
    patternInput.value = draft.pattern;
    ctx.syncControls();
    ctx.renderPreview();
  };

  const onPatternListClick = (e: Event): void => {
    const target = (e.target as HTMLElement | null)?.closest<HTMLButtonElement>(
      '[data-fc-pattern]',
    );
    if (!target) return;
    const pattern = target.dataset.fcPattern;
    if (!pattern) return;
    touch('numFmt');
    draft.pattern = pattern;
    patternInput.value = pattern;
    ctx.syncControls();
    syncPatternPresetOptions();
    ctx.renderPreview();
  };

  on(catList, 'click', onCatClick as EventListener);
  on(catList, 'keydown', onCatKeyDown as EventListener);
  on(decimalsInput, 'input', onDecimalsInput);
  on(thousandsCk.input, 'change', onThousandsChange);
  on(negativeOptions, 'click', onNegativeStyleClick as EventListener);
  on(symbolSelect, 'change', onSymbolChange);
  on(patternInput, 'input', onPatternInput);
  on(patternPresetSelect, 'change', onPatternPresetChange);
  on(patternList, 'click', onPatternListClick);

  return { sync, syncMixed };
}
