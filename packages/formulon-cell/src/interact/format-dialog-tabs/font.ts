// Font tab DOM for the Format Cells dialog. Bold/italic/underline/strike
// checkboxes plus family/size/color pickers and swatch grids.

import type { Locale, Strings } from '../../i18n/strings.js';
import {
  appendDialogDatalistOptions,
  createDialogSelect,
} from '../../toolbar/dialogs/form-controls.js';
import { fontFamilyCatalog, fontFamilyOptions } from '../../toolbar/ribbon/font-family-options.js';
import { FONT_SIZES } from '../../toolbar/ribbon-model.js';
import { appendDialogOptionButton, createDialogToggleButton } from '../dialog-shell.js';
import { makeButton, makeCheckbox, makeSwatches } from '../format-dialog-dom.js';

interface RevealSchedule {
  generation: number;
  frameId: number | null;
  scheduled: boolean;
}

const isRevealable = (list: HTMLElement): boolean =>
  list.isConnected && !list.closest('[hidden],[aria-hidden="true"]');

/**
 * Hydration runs before the dialog shell is opened.  Calling scrollIntoView
 * while the overlay or its tab panel is hidden is ignored by browsers, so a
 * selected family/size row would stay outside the list viewport on first
 * open.  Visible refreshes still scroll synchronously; hidden refreshes wait
 * for the owner window's next frame and re-check the live DOM.
 */
const revealSelected = (
  list: HTMLElement,
  schedule: RevealSchedule,
  ownerWindow: Window | null,
): void => {
  const revealNow = (): void => {
    if (!isRevealable(list)) return;
    list
      .querySelector<HTMLElement>('[aria-selected="true"]')
      ?.scrollIntoView?.({ block: 'nearest' });
  };

  if (isRevealable(list)) {
    schedule.generation += 1;
    schedule.scheduled = false;
    if (schedule.frameId !== null) {
      ownerWindow?.cancelAnimationFrame?.(schedule.frameId);
      schedule.frameId = null;
    }
    revealNow();
    return;
  }
  if (!list.isConnected || schedule.scheduled) return;

  schedule.scheduled = true;
  const generation = ++schedule.generation;
  const run = (): void => {
    if (!schedule.scheduled || schedule.generation !== generation) return;
    schedule.scheduled = false;
    schedule.frameId = null;
    revealNow();
  };
  const requestFrame = ownerWindow?.requestAnimationFrame?.bind(ownerWindow);
  if (requestFrame) {
    schedule.frameId = requestFrame(run);
  } else if (typeof queueMicrotask === 'function') {
    queueMicrotask(run);
  } else {
    void Promise.resolve().then(run);
  }
};

export interface FontTabRefs {
  boldCk: ReturnType<typeof makeCheckbox>;
  italicCk: ReturnType<typeof makeCheckbox>;
  underlineSelect: HTMLSelectElement;
  strikeCk: ReturnType<typeof makeCheckbox>;
  superscriptCk: ReturnType<typeof makeCheckbox>;
  subscriptCk: ReturnType<typeof makeCheckbox>;
  /** Mac-only effects; the controller hides them on other platforms. */
  outlineCk: ReturnType<typeof makeCheckbox>;
  shadowCk: ReturnType<typeof makeCheckbox>;
  normalFontCk: ReturnType<typeof makeCheckbox>;
  fontStyleList: HTMLDivElement;
  familyInput: HTMLInputElement;
  sizeInput: HTMLInputElement;
  colorInput: HTMLInputElement;
  colorReset: HTMLButtonElement;
  fontSwatches: ReturnType<typeof makeSwatches>;
  fontSwatchesToggle: HTMLButtonElement;
  fontSwatchesFlyout: HTMLDivElement;
  fontPreviewBox: HTMLDivElement;
  fontLocale?: Locale;
  syncFontFamilyOptions: (current: string, selectFirstWhenEmpty?: boolean) => void;
  syncFontSizeOptions: (current: number | undefined) => void;
}

export function createFontTab(
  panel: HTMLDivElement,
  t: Strings['formatDialog'],
  locale: Locale,
  sourceFamilies?: readonly string[],
): FontTabRefs {
  const ownerWindow = panel.ownerDocument.defaultView;
  const familySource = sourceFamilies ? fontFamilyCatalog(locale, sourceFamilies) : undefined;
  const familyRevealSchedule: RevealSchedule = {
    generation: 0,
    frameId: null,
    scheduled: false,
  };
  const sizeRevealSchedule: RevealSchedule = {
    generation: 0,
    frameId: null,
    scheduled: false,
  };
  const styleRow = document.createElement('div');
  styleRow.className = 'fc-fmtdlg__choice-grid fc-fmtdlg__font-effects';
  panel.appendChild(styleRow);

  const boldCk = makeCheckbox(t.fontBold);
  boldCk.input.dataset.fcCheck = 'bold';
  const italicCk = makeCheckbox(t.fontItalic);
  italicCk.input.dataset.fcCheck = 'italic';
  const underlineLabel = document.createElement('label');
  underlineLabel.className = 'fc-fmtdlg__font-underline';
  underlineLabel.textContent = t.fontUnderline;
  const underlineSelect = createDialogSelect(
    [
      { value: '', label: t.fontUnderlineNone },
      { value: 'single', label: t.fontUnderlineSingle },
      { value: 'double', label: t.fontUnderlineDouble },
      { value: 'singleAccounting', label: t.fontUnderlineSingleAccounting },
      { value: 'doubleAccounting', label: t.fontUnderlineDoubleAccounting },
    ],
    '',
    { ariaLabel: t.fontUnderline },
  );
  underlineSelect.dataset.fcInput = 'underline';
  underlineLabel.appendChild(underlineSelect);
  const strikeCk = makeCheckbox(t.fontStrike);
  strikeCk.input.dataset.fcCheck = 'strike';
  const superscriptCk = makeCheckbox(t.fontSuperscript);
  superscriptCk.input.dataset.fcCheck = 'superscript';
  const subscriptCk = makeCheckbox(t.fontSubscript);
  subscriptCk.input.dataset.fcCheck = 'subscript';
  const outlineCk = makeCheckbox(t.fontOutline);
  outlineCk.input.dataset.fcCheck = 'fontOutline';
  const shadowCk = makeCheckbox(t.fontShadow);
  shadowCk.input.dataset.fcCheck = 'fontShadow';
  styleRow.append(
    boldCk.wrap,
    italicCk.wrap,
    underlineLabel,
    strikeCk.wrap,
    superscriptCk.wrap,
    subscriptCk.wrap,
    outlineCk.wrap,
    shadowCk.wrap,
  );

  const normalFontCk = makeCheckbox(t.normalFont);
  normalFontCk.input.dataset.fcCheck = 'normalFont';
  normalFontCk.wrap.classList.add('fc-fmtdlg__normal-font');
  panel.appendChild(normalFontCk.wrap);

  // Font family
  const familyRow = document.createElement('label');
  familyRow.className = 'fc-fmtdlg__row fc-fmtdlg__font-family-row';
  const familyLabel = document.createElement('span');
  familyLabel.textContent = t.fontFamily;
  const familyInput = document.createElement('input');
  familyInput.type = 'text';
  familyInput.setAttribute('aria-label', t.fontFamily);
  familyInput.dataset.fcInput = 'family';
  familyInput.spellcheck = false;
  familyInput.autocomplete = 'off';
  const familyListId = `fc-fmtdlg-fonts-${Math.random().toString(36).slice(2, 8)}`;
  familyInput.setAttribute('list', familyListId);
  const familyDatalist = document.createElement('datalist');
  familyDatalist.id = familyListId;
  familyRow.append(familyLabel, familyInput, familyDatalist);
  panel.appendChild(familyRow);

  const familyList = document.createElement('div');
  familyList.className = 'fc-fmtdlg__font-list fc-fmtdlg__font-list--family';
  familyList.setAttribute('role', 'listbox');
  familyList.setAttribute('aria-label', t.fontFamily);
  panel.appendChild(familyList);

  const syncFontFamilyOptions = (current: string, selectFirstWhenEmpty = true): void => {
    const families = fontFamilyOptions(locale, current, familySource);
    familyDatalist.replaceChildren();
    appendDialogDatalistOptions(
      familyDatalist,
      families.map(({ value, label }) => ({ value, label })),
    );
    familyList.replaceChildren();
    for (const [index, family] of families.entries()) {
      const item = appendDialogOptionButton(familyList, {
        label: family.label,
        baseClass: 'fc-fmtdlg__font-list-item',
        datasetKey: 'fcFontFamily',
        value: family.value,
        selected: family.value === current || (selectFirstWhenEmpty && !current && index === 0),
      });
      if (family.role) item.dataset.fcFontRole = family.role;
      item.addEventListener('click', () => {
        familyInput.value = family.label;
        familyInput.dispatchEvent(
          new CustomEvent('fc-font-family-option', {
            bubbles: true,
            detail: { value: family.value },
          }),
        );
      });
    }
    revealSelected(familyList, familyRevealSchedule, ownerWindow);
  };
  syncFontFamilyOptions('');

  // The style column has no free-text input of its own, so it carries a bare
  // caption to keep the three column headings on one line.
  const fontStyleLabel = document.createElement('div');
  fontStyleLabel.className = 'fc-fmtdlg__font-style-label';
  fontStyleLabel.textContent = t.fontStyle;
  panel.appendChild(fontStyleLabel);

  const fontStyleList = document.createElement('div');
  fontStyleList.className = 'fc-fmtdlg__font-list fc-fmtdlg__font-list--style';
  fontStyleList.setAttribute('role', 'listbox');
  fontStyleList.setAttribute('aria-label', t.fontStyle);
  const fontStyleOptions = [
    { id: 'regular', label: t.fontRegular },
    { id: 'italic', label: t.fontItalic },
    { id: 'bold', label: t.fontBold },
    { id: 'boldItalic', label: `${t.fontBold} ${t.fontItalic}` },
  ] as const;
  for (const [index, option] of fontStyleOptions.entries()) {
    appendDialogOptionButton(fontStyleList, {
      label: option.label,
      baseClass: 'fc-fmtdlg__font-list-item',
      datasetKey: 'fcFontStyle',
      value: option.id,
      selected: index === 0,
    });
  }
  panel.appendChild(fontStyleList);

  // Font size
  const sizeRow = document.createElement('label');
  sizeRow.className = 'fc-fmtdlg__row fc-fmtdlg__font-size-row';
  const sizeLabel = document.createElement('span');
  sizeLabel.textContent = t.fontSize;
  const sizeInput = document.createElement('input');
  sizeInput.type = 'number';
  sizeInput.setAttribute('aria-label', t.fontSize);
  sizeInput.dataset.fcInput = 'size';
  sizeInput.min = '1';
  sizeInput.max = '409';
  sizeInput.step = '1';
  sizeRow.append(sizeLabel, sizeInput);
  panel.appendChild(sizeRow);

  const sizeList = document.createElement('div');
  sizeList.className = 'fc-fmtdlg__font-list fc-fmtdlg__font-list--size';
  sizeList.setAttribute('role', 'listbox');
  sizeList.setAttribute('aria-label', t.fontSize);
  for (const size of FONT_SIZES) {
    const item = appendDialogOptionButton(sizeList, {
      label: String(size),
      baseClass: 'fc-fmtdlg__font-list-item',
      datasetKey: 'fcFontSize',
      value: String(size),
      selected: false,
    });
    item.addEventListener('click', () => {
      sizeInput.value = String(size);
      sizeInput.dispatchEvent(new Event('input', { bubbles: true }));
    });
  }
  panel.appendChild(sizeList);
  const syncFontSizeOptions = (current: number | undefined): void => {
    for (const item of sizeList.querySelectorAll<HTMLButtonElement>('[data-fc-font-size]')) {
      const isSelected = current !== undefined && Number(item.dataset.fcFontSize) === current;
      item.setAttribute('aria-selected', isSelected ? 'true' : 'false');
    }
    revealSelected(sizeList, sizeRevealSchedule, ownerWindow);
  };

  // Both the dialog overlay and this tab panel use `hidden` during
  // construction/tab switches. Observe those transitions so a list that was
  // hydrated while another tab was active is also revealed when the user
  // later opens the Font tab.
  if (typeof MutationObserver !== 'undefined') {
    const visibilityObserver = new MutationObserver(() => {
      if (!panel.isConnected) {
        visibilityObserver.disconnect();
        return;
      }
      revealSelected(familyList, familyRevealSchedule, ownerWindow);
      revealSelected(sizeList, sizeRevealSchedule, ownerWindow);
    });
    const dialogRoot = panel.closest<HTMLElement>('[role="dialog"]');
    for (const target of [dialogRoot, panel]) {
      target?.isConnected &&
        visibilityObserver.observe(target, {
          attributes: true,
          attributeFilter: ['hidden', 'aria-hidden'],
        });
    }
  }

  // Font color. The swatch palette is taller than the space this tab has left,
  // so it hangs off the color control as a flyout instead of sitting inline the
  // way the fill tab's palette does.
  const colorRow = document.createElement('div');
  colorRow.className = 'fc-fmtdlg__row fc-fmtdlg__font-color-row';
  const colorLabel = document.createElement('span');
  colorLabel.textContent = t.color;
  const colorInput = document.createElement('input');
  colorInput.type = 'color';
  colorInput.setAttribute('aria-label', t.color);
  colorInput.dataset.fcColor = 'font';
  const fontSwatchesToggle = createDialogToggleButton({
    label: t.themeColors,
    baseClass: 'fc-fmtdlg__color-toggle',
    title: t.themeColors,
    popup: true,
  });
  const fontSwatchesChevron = document.createElement('span');
  fontSwatchesChevron.className = 'fc-fmtdlg__color-toggle-chevron';
  fontSwatchesChevron.setAttribute('aria-hidden', 'true');
  fontSwatchesToggle.appendChild(fontSwatchesChevron);
  const colorControl = document.createElement('span');
  colorControl.className = 'fc-fmtdlg__color-control';
  colorControl.append(colorInput, fontSwatchesToggle);
  const colorReset = makeButton(t.resetToDefault);
  colorRow.append(colorLabel, colorControl, colorReset);
  panel.appendChild(colorRow);
  const fontSwatches = makeSwatches('font', t.themeColors, t.standardColors);
  const fontSwatchesFlyout = document.createElement('div');
  fontSwatchesFlyout.className = 'fc-fmtdlg__color-flyout';
  fontSwatchesFlyout.hidden = true;
  fontSwatchesFlyout.appendChild(fontSwatches.el);
  panel.appendChild(fontSwatchesFlyout);

  const fontPreview = document.createElement('div');
  fontPreview.className = 'fc-fmtdlg__font-preview';
  const fontPreviewLabel = document.createElement('div');
  fontPreviewLabel.className = 'fc-fmtdlg__font-preview-label';
  fontPreviewLabel.textContent = t.preview;
  const fontPreviewBox = document.createElement('div');
  fontPreviewBox.className = 'fc-fmtdlg__font-preview-box';
  const fontPreviewRun = document.createElement('span');
  fontPreviewRun.className = 'fc-fmtdlg__font-preview-run';
  fontPreviewRun.textContent = t.previewText;
  fontPreviewBox.appendChild(fontPreviewRun);
  fontPreview.append(fontPreviewLabel, fontPreviewBox);
  panel.appendChild(fontPreview);

  return {
    boldCk,
    italicCk,
    underlineSelect,
    strikeCk,
    superscriptCk,
    subscriptCk,
    outlineCk,
    shadowCk,
    normalFontCk,
    fontStyleList,
    familyInput,
    sizeInput,
    colorInput,
    colorReset,
    fontSwatches,
    fontSwatchesToggle,
    fontSwatchesFlyout,
    fontPreviewBox,
    fontLocale: locale,
    syncFontFamilyOptions,
    syncFontSizeOptions,
  };
}
