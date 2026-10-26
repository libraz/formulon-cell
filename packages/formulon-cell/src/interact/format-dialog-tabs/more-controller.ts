// "More" tab behaviour for the Format Cells dialog: hyperlink, comment and
// the data-validation editor, which the Data Validation dialog mode shows on
// its own. Date/time validation bounds round-trip through native pickers.

import { coerceInput, parseLocalizedNumericInput } from '../../commands/coerce-input.js';
import { formatEditableNumber } from '../../engine/value.js';
import { excelSerialToIsoDate } from '../../format/date-serial.js';
import type { ValidationErrorStyle, ValidationOp } from '../../store/store.js';
import type { ValidationKind } from '../format-dialog-model.js';
import { type FormatTabContext, type FormatTabController, markMixed } from './controller.js';
import type { MoreTabRefs } from './more.js';

/** Convert a spreadsheet date serial to a native `<input type="date">` value
 *  (`yyyy-mm-dd`, UTC). Returns '' for non-finite serials. */
const serialToDateInputValue = (serial: number, date1904: boolean): string => {
  return excelSerialToIsoDate(serial, date1904) ?? '';
};

/** Convert a day-fraction time serial to a native `<input type="time">` value
 *  (`HH:mm`, or `HH:mm:ss` when the serial carries seconds). */
const serialToTimeInputValue = (serial: number): string => {
  if (!Number.isFinite(serial)) return '';
  let total = Math.round((serial % 1) * 86_400);
  total = ((total % 86_400) + 86_400) % 86_400;
  const hh = String(Math.floor(total / 3600)).padStart(2, '0');
  const mm = String(Math.floor((total % 3600) / 60)).padStart(2, '0');
  const ss = total % 60;
  return ss ? `${hh}:${mm}:${String(ss).padStart(2, '0')}` : `${hh}:${mm}`;
};

/** Format a stored bound (serial or plain number) for the bound `<input>` value,
 *  matching the input type chosen for the validation kind. */
const boundInputValue = (
  kind: ValidationKind,
  value: number,
  locale: string | undefined,
  date1904: boolean,
): string => {
  if (kind === 'date') return serialToDateInputValue(value, date1904);
  if (kind === 'time') return serialToTimeInputValue(value);
  return formatEditableNumber(value, locale);
};

/** Parse a bound `<input>` value back into a stored number. Date/time kinds
 *  route the string through `coerceInput` so `yyyy-mm-dd` / `HH:mm` become the
 *  matching spreadsheet serial; other kinds parse a plain number. Returns null
 *  when the field is empty or unparseable so the previous bound is kept. */
const parseBoundInputValue = (
  kind: ValidationKind,
  raw: string,
  locale: string | undefined,
  date1904: boolean,
): number | null => {
  if (kind === 'date' || kind === 'time') {
    const coerced = coerceInput(raw, { date1904 });
    return coerced.kind === 'number' ? coerced.value : null;
  }
  const n = parseLocalizedNumericInput(raw, locale);
  if (n === null) return null;
  if ((kind === 'whole' || kind === 'textLength') && !Number.isInteger(n)) return null;
  return n;
};

export function attachMoreTab(ctx: FormatTabContext, refs: MoreTabRefs): FormatTabController {
  const { draft, on, touch, isMixed, syncControls } = ctx;
  const getLocale = ctx.getLocale;
  const getDate1904 = ctx.getDate1904;
  const {
    hlInput,
    hlClear,
    commentArea,
    commentClear,
    validationKindSelect,
    validationSection,
    validationOpRow,
    validationOpSelect,
    validationARow,
    validationAInput,
    validationBRow,
    validationBInput,
    validationFormulaRow,
    validationFormulaInput,
    validationListSourceKindRow,
    validationListLiteralRadio,
    validationListRangeRadio,
    validationRow,
    validationArea,
    validationClear,
    validationListRangeRow,
    validationListRangeInput,
    validationShowDropdownRow,
    validationShowDropdownInput,
    validationAllowBlankRow,
    validationAllowBlankInput,
    validationErrorStyleRow,
    validationErrorStyleSelect,
    validationShowInputMessageRow,
    validationShowInputMessageInput,
    validationPromptTitleRow,
    validationPromptTitleInput,
    validationPromptMessageRow,
    validationPromptMessageArea,
    validationShowErrorMessageRow,
    validationShowErrorMessageInput,
    validationErrorTitleRow,
    validationErrorTitleInput,
    validationErrorMessageRow,
    validationErrorMessageArea,
  } = refs;

  const dialog = validationSection.closest<HTMLElement>('[role="dialog"]');
  const invalidBoundInputs = new Set<HTMLInputElement>();
  const setBoundInputValidity = (input: HTMLInputElement, valid: boolean): void => {
    if (valid) {
      invalidBoundInputs.delete(input);
      delete input.dataset.fcInvalid;
      input.removeAttribute('aria-invalid');
      input.setCustomValidity('');
      return;
    }
    invalidBoundInputs.add(input);
    input.dataset.fcInvalid = 'true';
    input.setAttribute('aria-invalid', 'true');
    input.setCustomValidity('invalid');
  };

  /** Switch the A/B bound inputs to a native date/time picker for date/time
   *  validation (so bounds are pickable instead of raw serials) and back to a
   *  localized text field for numeric validation. */
  const applyBoundInputMode = (kind: ValidationKind): void => {
    const type = kind === 'date' ? 'date' : kind === 'time' ? 'time' : 'text';
    for (const input of [validationAInput, validationBInput]) {
      if (input.type !== type) input.type = type;
      if (type === 'time') input.step = '1';
      else input.removeAttribute('step');
      if (type === 'text') input.inputMode = kind === 'decimal' ? 'decimal' : 'numeric';
      else input.removeAttribute('inputmode');
    }
  };

  const syncValidationVisibility = (): void => {
    const k = draft.validationKind;
    applyBoundInputMode(k);
    const isBounded =
      k === 'whole' || k === 'decimal' || k === 'date' || k === 'time' || k === 'textLength';
    const isListLike = k === 'list';
    const isCustom = k === 'custom';
    const isActive = k !== 'none';
    validationOpRow.hidden = !isBounded;
    validationARow.hidden = !isBounded;
    validationBRow.hidden =
      !isBounded || (draft.validationOp !== 'between' && draft.validationOp !== 'notBetween');
    validationFormulaRow.hidden = !isCustom;
    validationListSourceKindRow.hidden = !isListLike;
    validationRow.hidden = !isListLike || draft.validationListSourceKind !== 'literal';
    validationListRangeRow.hidden = !isListLike || draft.validationListSourceKind !== 'range';
    validationShowDropdownRow.hidden = !isListLike;
    validationAllowBlankRow.hidden = !isActive;
    validationErrorStyleRow.hidden = !isActive;
    validationShowInputMessageRow.hidden = !isActive;
    validationPromptTitleRow.hidden = !isActive || !draft.validationShowInputMessage;
    validationPromptMessageRow.hidden = !isActive || !draft.validationShowInputMessage;
    validationShowErrorMessageRow.hidden = !isActive;
    validationErrorTitleRow.hidden = !isActive || !draft.validationShowErrorMessage;
    validationErrorMessageRow.hidden = !isActive || !draft.validationShowErrorMessage;
  };

  const sync = (): void => {
    // `sync()` also runs while FormatDialog is hydrating a new open. Drop an
    // invalid edit left by a previous open before rehydrating the stored draft.
    if (dialog?.hidden) {
      setBoundInputValidity(validationAInput, true);
      setBoundInputValidity(validationBInput, true);
    }
    hlInput.value = draft.hyperlink;
    commentArea.value = draft.comment;
    validationArea.value = draft.validationList;
    validationListRangeInput.value = draft.validationListRange;
    validationListLiteralRadio.input.checked = draft.validationListSourceKind === 'literal';
    validationListRangeRadio.input.checked = draft.validationListSourceKind === 'range';
    validationShowDropdownInput.checked = draft.validationShowDropdown;
    validationKindSelect.value = draft.validationKind;
    validationOpSelect.value = draft.validationOp;
    applyBoundInputMode(draft.validationKind);
    if (!invalidBoundInputs.has(validationAInput)) {
      validationAInput.value = boundInputValue(
        draft.validationKind,
        draft.validationA,
        getLocale(),
        getDate1904(),
      );
    }
    if (!invalidBoundInputs.has(validationBInput)) {
      validationBInput.value = boundInputValue(
        draft.validationKind,
        draft.validationB,
        getLocale(),
        getDate1904(),
      );
    }
    validationFormulaInput.value = draft.validationFormula;
    validationAllowBlankInput.checked = draft.validationAllowBlank;
    validationErrorStyleSelect.value = draft.validationErrorStyle;
    validationShowInputMessageInput.checked = draft.validationShowInputMessage;
    validationPromptTitleInput.value = draft.validationPromptTitle;
    validationPromptMessageArea.value = draft.validationPromptMessage;
    validationShowErrorMessageInput.checked = draft.validationShowErrorMessage;
    validationErrorTitleInput.value = draft.validationErrorTitle;
    validationErrorMessageArea.value = draft.validationErrorMessage;
    syncValidationVisibility();
  };

  const syncMixed = (): void => {
    if (isMixed('hyperlink')) {
      hlInput.value = '';
      markMixed(hlInput, true, 'hyperlink');
    }
    if (isMixed('comment')) {
      commentArea.value = '';
      markMixed(commentArea, true, 'comment');
    }
    if (isMixed('validation')) {
      validationKindSelect.value = '';
      validationOpSelect.value = '';
      validationAInput.value = '';
      validationBInput.value = '';
      validationFormulaInput.value = '';
      validationListSourceKindRow.querySelectorAll<HTMLInputElement>('input').forEach((input) => {
        input.checked = false;
        input.indeterminate = false;
      });
      validationArea.value = '';
      validationListRangeInput.value = '';
      validationShowDropdownInput.checked = false;
      validationShowDropdownInput.indeterminate = true;
      validationAllowBlankInput.checked = false;
      validationAllowBlankInput.indeterminate = true;
      validationErrorStyleSelect.value = '';
      validationShowInputMessageInput.checked = false;
      validationShowInputMessageInput.indeterminate = true;
      validationPromptTitleInput.value = '';
      validationPromptMessageArea.value = '';
      validationShowErrorMessageInput.checked = false;
      validationShowErrorMessageInput.indeterminate = true;
      validationErrorTitleInput.value = '';
      validationErrorMessageArea.value = '';
      for (const control of [
        validationKindSelect,
        validationOpSelect,
        validationAInput,
        validationBInput,
        validationFormulaInput,
        validationArea,
        validationListRangeInput,
        validationShowDropdownInput,
        validationAllowBlankInput,
        validationErrorStyleSelect,
        validationShowInputMessageInput,
        validationPromptTitleInput,
        validationPromptMessageArea,
        validationShowErrorMessageInput,
        validationErrorTitleInput,
        validationErrorMessageArea,
      ]) {
        markMixed(control, true, 'validation');
      }
      validationListSourceKindRow.querySelectorAll<HTMLInputElement>('input').forEach((input) => {
        markMixed(input, true, 'validation');
      });
      markMixed(validationKindSelect, true, 'validation');
    }
  };

  // ── Events ─────────────────────────────────────────────────────────────
  const onHlInput = (): void => {
    touch('hyperlink');
    draft.hyperlink = hlInput.value;
  };
  const onHlClear = (): void => {
    touch('hyperlink');
    draft.hyperlink = '';
    hlInput.value = '';
  };
  const onCommentInput = (): void => {
    touch('comment');
    draft.comment = commentArea.value;
  };
  const onCommentClear = (): void => {
    touch('comment');
    draft.comment = '';
    commentArea.value = '';
  };
  const onValidationInput = (): void => {
    touch('validation');
    draft.validationList = validationArea.value;
    syncControls();
  };
  const onValidationClear = (): void => {
    touch('validation');
    draft.validationList = '';
    validationArea.value = '';
    syncControls();
  };
  const onValidationListRangeInput = (): void => {
    touch('validation');
    draft.validationListRange = validationListRangeInput.value;
    syncControls();
  };
  const onValidationListSourceKindChange = (): void => {
    touch('validation');
    if (validationListLiteralRadio.input.checked) draft.validationListSourceKind = 'literal';
    else if (validationListRangeRadio.input.checked) draft.validationListSourceKind = 'range';
    syncControls();
  };
  const onValidationShowDropdownChange = (): void => {
    touch('validation');
    draft.validationShowDropdown = validationShowDropdownInput.checked;
    syncControls();
  };
  const onValidationKindChange = (): void => {
    touch('validation');
    setBoundInputValidity(validationAInput, true);
    setBoundInputValidity(validationBInput, true);
    draft.validationKind = validationKindSelect.value as ValidationKind;
    // Switching between numeric / date / time kinds swaps the bound-input type,
    // so re-render the stored bounds in the new type's value format.
    applyBoundInputMode(draft.validationKind);
    validationAInput.value = boundInputValue(
      draft.validationKind,
      draft.validationA,
      getLocale(),
      getDate1904(),
    );
    validationBInput.value = boundInputValue(
      draft.validationKind,
      draft.validationB,
      getLocale(),
      getDate1904(),
    );
    syncControls();
  };
  const onValidationOpChange = (): void => {
    touch('validation');
    draft.validationOp = validationOpSelect.value as ValidationOp;
    if (draft.validationOp !== 'between' && draft.validationOp !== 'notBetween') {
      // A hidden upper bound is no longer part of the rule. Discard an
      // unfinished edit there so it cannot block a valid one-sided commit.
      setBoundInputValidity(validationBInput, true);
    }
    syncControls();
  };
  const onValidationAInput = (): void => {
    touch('validation');
    const n = parseBoundInputValue(
      draft.validationKind,
      validationAInput.value,
      getLocale(),
      getDate1904(),
    );
    setBoundInputValidity(validationAInput, n !== null);
    if (n !== null) {
      draft.validationA = n;
      draft.validationAFormula = undefined;
      syncControls();
    } else {
      syncValidationVisibility();
    }
  };
  const onValidationBInput = (): void => {
    touch('validation');
    const n = parseBoundInputValue(
      draft.validationKind,
      validationBInput.value,
      getLocale(),
      getDate1904(),
    );
    setBoundInputValidity(validationBInput, n !== null);
    if (n !== null) {
      draft.validationB = n;
      draft.validationBFormula = undefined;
      syncControls();
    } else {
      syncValidationVisibility();
    }
  };
  const onValidationFormulaInput = (): void => {
    touch('validation');
    draft.validationFormula = validationFormulaInput.value;
    syncControls();
  };
  const onValidationAllowBlankChange = (): void => {
    touch('validation');
    draft.validationAllowBlank = validationAllowBlankInput.checked;
    syncControls();
  };
  const onValidationErrorStyleChange = (): void => {
    touch('validation');
    draft.validationErrorStyle = validationErrorStyleSelect.value as ValidationErrorStyle;
    syncControls();
  };
  const onValidationShowInputMessageChange = (): void => {
    touch('validation');
    draft.validationShowInputMessage = validationShowInputMessageInput.checked;
    syncControls();
  };
  const onValidationPromptTitleInput = (): void => {
    touch('validation');
    draft.validationPromptTitle = validationPromptTitleInput.value;
    syncControls();
  };
  const onValidationPromptMessageInput = (): void => {
    touch('validation');
    draft.validationPromptMessage = validationPromptMessageArea.value;
    syncControls();
  };
  const onValidationShowErrorMessageChange = (): void => {
    touch('validation');
    draft.validationShowErrorMessage = validationShowErrorMessageInput.checked;
    syncControls();
  };
  const onValidationErrorTitleInput = (): void => {
    touch('validation');
    draft.validationErrorTitle = validationErrorTitleInput.value;
    syncControls();
  };
  const onValidationErrorMessageInput = (): void => {
    touch('validation');
    draft.validationErrorMessage = validationErrorMessageArea.value;
    syncControls();
  };

  on(hlInput, 'input', onHlInput);
  on(hlClear, 'click', onHlClear);
  on(commentArea, 'input', onCommentInput);
  on(commentClear, 'click', onCommentClear);
  on(validationArea, 'input', onValidationInput);
  on(validationClear, 'click', onValidationClear);
  on(validationListRangeInput, 'input', onValidationListRangeInput);
  on(validationListLiteralRadio.input, 'change', onValidationListSourceKindChange);
  on(validationListRangeRadio.input, 'change', onValidationListSourceKindChange);
  on(validationShowDropdownInput, 'change', onValidationShowDropdownChange);
  on(validationKindSelect, 'change', onValidationKindChange);
  on(validationOpSelect, 'change', onValidationOpChange);
  on(validationAInput, 'input', onValidationAInput);
  on(validationBInput, 'input', onValidationBInput);
  on(validationFormulaInput, 'input', onValidationFormulaInput);
  on(validationAllowBlankInput, 'change', onValidationAllowBlankChange);
  on(validationErrorStyleSelect, 'change', onValidationErrorStyleChange);
  on(validationShowInputMessageInput, 'change', onValidationShowInputMessageChange);
  on(validationPromptTitleInput, 'input', onValidationPromptTitleInput);
  on(validationPromptMessageArea, 'input', onValidationPromptMessageInput);
  on(validationShowErrorMessageInput, 'change', onValidationShowErrorMessageChange);
  on(validationErrorTitleInput, 'input', onValidationErrorTitleInput);
  on(validationErrorMessageArea, 'input', onValidationErrorMessageInput);

  // FormatDialog's OK listener is registered after this controller is attached.
  // Stop that listener while a numeric/date/time bound still carries an
  // invalid edit, rather than letting the untouched draft value win silently.
  const okButton = dialog?.querySelector<HTMLButtonElement>(
    '.fc-fmtdlg__footer .fc-fmtdlg__btn--primary',
  );
  if (okButton) {
    on(okButton, 'click', (event) => {
      const invalid = [...invalidBoundInputs][0];
      if (!invalid) return;
      event.preventDefault();
      event.stopImmediatePropagation();
      invalid.focus();
    });
  }

  return { sync, syncMixed };
}
