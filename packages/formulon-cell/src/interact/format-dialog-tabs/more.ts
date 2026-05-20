// "More" tab DOM for the Format Cells dialog — hyperlink, comment, and the
// full data-validation editor (kind/op/values/list-source/error style/
// input + error messages).

import type { Strings } from '../../i18n/strings.js';
import type { ValidationOp } from '../../store/store.js';
import { createDialogSelect } from '../../toolbar/dialogs/form-controls.js';
import { makeButton, makeListSourceRadio, makeSection } from '../format-dialog-dom.js';
import type { ValidationKind } from '../format-dialog-model.js';

export interface MoreTabRefs {
  hyperlinkSection: HTMLDivElement;
  commentSection: HTMLDivElement;
  validationSection: HTMLDivElement;
  hlInput: HTMLInputElement;
  hlClear: HTMLButtonElement;
  commentArea: HTMLTextAreaElement;
  commentClear: HTMLButtonElement;
  validationKindSelect: HTMLSelectElement;
  validationOpRow: HTMLLabelElement;
  validationOpSelect: HTMLSelectElement;
  validationARow: HTMLLabelElement;
  validationAInput: HTMLInputElement;
  validationBRow: HTMLLabelElement;
  validationBInput: HTMLInputElement;
  validationFormulaRow: HTMLLabelElement;
  validationFormulaInput: HTMLInputElement;
  validationListSourceKindRow: HTMLDivElement;
  validationListLiteralRadio: ReturnType<typeof makeListSourceRadio>;
  validationListRangeRadio: ReturnType<typeof makeListSourceRadio>;
  validationRow: HTMLDivElement;
  validationArea: HTMLTextAreaElement;
  validationClear: HTMLButtonElement;
  validationListRangeRow: HTMLLabelElement;
  validationListRangeInput: HTMLInputElement;
  validationAllowBlankRow: HTMLLabelElement;
  validationAllowBlankInput: HTMLInputElement;
  validationErrorStyleRow: HTMLLabelElement;
  validationErrorStyleSelect: HTMLSelectElement;
  validationShowInputMessageRow: HTMLLabelElement;
  validationShowInputMessageInput: HTMLInputElement;
  validationPromptTitleRow: HTMLLabelElement;
  validationPromptTitleInput: HTMLInputElement;
  validationPromptMessageRow: HTMLDivElement;
  validationPromptMessageArea: HTMLTextAreaElement;
  validationShowErrorMessageRow: HTMLLabelElement;
  validationShowErrorMessageInput: HTMLInputElement;
  validationErrorTitleRow: HTMLLabelElement;
  validationErrorTitleInput: HTMLInputElement;
  validationErrorMessageRow: HTMLDivElement;
  validationErrorMessageArea: HTMLTextAreaElement;
}

export function createMoreTab(panel: HTMLDivElement, t: Strings['formatDialog']): MoreTabRefs {
  const hyperlinkSection = makeSection(t.hyperlink);
  const commentSection = makeSection(t.comment);
  const validationSection = makeSection(t.validationLegend);
  panel.append(hyperlinkSection, commentSection, validationSection);

  const hlRow = document.createElement('div');
  hlRow.className = 'fc-fmtdlg__row';
  hyperlinkSection.appendChild(hlRow);
  const hlLabel = document.createElement('span');
  hlLabel.textContent = t.hyperlink;
  const hlInput = document.createElement('input');
  hlInput.type = 'text';
  hlInput.setAttribute('aria-label', t.hyperlink);
  hlInput.spellcheck = false;
  hlInput.autocomplete = 'off';
  hlInput.placeholder = t.hyperlinkPlaceholder;
  const hlClear = makeButton(t.clearField);
  hlRow.append(hlLabel, hlInput, hlClear);

  const commentRow = document.createElement('div');
  commentRow.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
  commentSection.appendChild(commentRow);
  const commentArea = document.createElement('textarea');
  commentArea.className = 'fc-fmtdlg__textarea';
  commentArea.rows = 3;
  commentArea.setAttribute('aria-label', t.comment);
  commentArea.placeholder = t.commentPlaceholder;
  const commentClear = makeButton(t.clearField);
  commentRow.append(commentArea, commentClear);

  // Kind selector — drives the visibility of the bound/formula/list rows.
  const validationKindRow = document.createElement('label');
  validationKindRow.className = 'fc-fmtdlg__row';
  const validationKindLabel = document.createElement('span');
  validationKindLabel.textContent = t.validationKind;
  const kindDefs: { id: ValidationKind; label: string }[] = [
    { id: 'none', label: t.validationKindNone },
    { id: 'list', label: t.validationKindList },
    { id: 'whole', label: t.validationKindWhole },
    { id: 'decimal', label: t.validationKindDecimal },
    { id: 'date', label: t.validationKindDate },
    { id: 'time', label: t.validationKindTime },
    { id: 'textLength', label: t.validationKindTextLength },
    { id: 'custom', label: t.validationKindCustom },
  ];
  const validationKindSelect = createDialogSelect(
    kindDefs.map((item) => ({ value: item.id, label: item.label })),
    'none',
    { className: '', ariaLabel: t.validationKind },
  );
  validationKindRow.append(validationKindLabel, validationKindSelect);
  validationSection.appendChild(validationKindRow);

  // Op selector — visible for whole/decimal/date/time/textLength.
  const validationOpRow = document.createElement('label');
  validationOpRow.className = 'fc-fmtdlg__row';
  const validationOpLabel = document.createElement('span');
  validationOpLabel.textContent = t.validationOp;
  const opDefs: { id: ValidationOp; label: string }[] = [
    { id: 'between', label: t.validationOpBetween },
    { id: 'notBetween', label: t.validationOpNotBetween },
    { id: '=', label: t.validationOpEq },
    { id: '<>', label: t.validationOpNeq },
    { id: '<', label: t.validationOpLt },
    { id: '<=', label: t.validationOpLte },
    { id: '>', label: t.validationOpGt },
    { id: '>=', label: t.validationOpGte },
  ];
  const validationOpSelect = createDialogSelect(
    opDefs.map((item) => ({ value: item.id, label: item.label })),
    'between',
    { className: '', ariaLabel: t.validationOp },
  );
  validationOpRow.append(validationOpLabel, validationOpSelect);
  validationSection.appendChild(validationOpRow);

  const validationARow = document.createElement('label');
  validationARow.className = 'fc-fmtdlg__row';
  const validationALabel = document.createElement('span');
  validationALabel.textContent = t.validationValueA;
  const validationAInput = document.createElement('input');
  validationAInput.type = 'number';
  validationAInput.setAttribute('aria-label', t.validationValueA);
  validationAInput.step = 'any';
  validationARow.append(validationALabel, validationAInput);
  validationSection.appendChild(validationARow);

  const validationBRow = document.createElement('label');
  validationBRow.className = 'fc-fmtdlg__row';
  const validationBLabel = document.createElement('span');
  validationBLabel.textContent = t.validationValueB;
  const validationBInput = document.createElement('input');
  validationBInput.type = 'number';
  validationBInput.setAttribute('aria-label', t.validationValueB);
  validationBInput.step = 'any';
  validationBRow.append(validationBLabel, validationBInput);
  validationSection.appendChild(validationBRow);

  // Custom-kind formula.
  const validationFormulaRow = document.createElement('label');
  validationFormulaRow.className = 'fc-fmtdlg__row';
  const validationFormulaLabel = document.createElement('span');
  validationFormulaLabel.textContent = t.validationFormula;
  const validationFormulaInput = document.createElement('input');
  validationFormulaInput.type = 'text';
  validationFormulaInput.setAttribute('aria-label', t.validationFormula);
  validationFormulaInput.spellcheck = false;
  validationFormulaInput.autocomplete = 'off';
  validationFormulaInput.placeholder = t.validationFormulaPlaceholder;
  validationFormulaRow.append(validationFormulaLabel, validationFormulaInput);
  validationSection.appendChild(validationFormulaRow);

  // List source — visible only when kind === 'list'. The radio toggles between
  // an inline value list (textarea) and a range reference (single-line input).
  const validationListSourceKindRow = document.createElement('div');
  validationListSourceKindRow.className = 'fc-fmtdlg__row fc-fmtdlg__row--inline';
  const validationListSourceKindLabel = document.createElement('span');
  validationListSourceKindLabel.textContent = t.validationListSourceKind;
  validationListSourceKindRow.appendChild(validationListSourceKindLabel);
  const validationListLiteralRadio = makeListSourceRadio('literal', t.validationListSourceLiteral);
  const validationListRangeRadio = makeListSourceRadio('range', t.validationListSourceRange);
  validationListSourceKindRow.append(
    validationListLiteralRadio.wrap,
    validationListRangeRadio.wrap,
  );
  validationSection.appendChild(validationListSourceKindRow);

  const validationRow = document.createElement('div');
  validationRow.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
  const validationArea = document.createElement('textarea');
  validationArea.className = 'fc-fmtdlg__textarea';
  validationArea.rows = 4;
  validationArea.setAttribute('aria-label', t.validationListSourceLiteral);
  validationArea.placeholder = t.validationListPlaceholder;
  const validationClear = makeButton(t.clearField);
  validationRow.append(validationArea, validationClear);
  validationSection.appendChild(validationRow);

  // Range-ref input. Hidden unless source kind === 'range'.
  const validationListRangeRow = document.createElement('label');
  validationListRangeRow.className = 'fc-fmtdlg__row';
  const validationListRangeLabel = document.createElement('span');
  validationListRangeLabel.textContent = t.validationListSourceRange;
  const validationListRangeInput = document.createElement('input');
  validationListRangeInput.type = 'text';
  validationListRangeInput.setAttribute('aria-label', t.validationListSourceRange);
  validationListRangeInput.spellcheck = false;
  validationListRangeInput.autocomplete = 'off';
  validationListRangeInput.placeholder = t.validationListRangePlaceholder;
  validationListRangeRow.append(validationListRangeLabel, validationListRangeInput);
  validationSection.appendChild(validationListRangeRow);

  // Allow-blank checkbox + error-style selector.
  const validationAllowBlankRow = document.createElement('label');
  validationAllowBlankRow.className = 'fc-fmtdlg__check';
  const validationAllowBlankInput = document.createElement('input');
  validationAllowBlankInput.type = 'checkbox';
  const validationAllowBlankSpan = document.createElement('span');
  validationAllowBlankSpan.textContent = t.validationAllowBlank;
  validationAllowBlankRow.append(validationAllowBlankInput, validationAllowBlankSpan);
  validationSection.appendChild(validationAllowBlankRow);

  const validationErrorStyleRow = document.createElement('label');
  validationErrorStyleRow.className = 'fc-fmtdlg__row';
  const validationErrorStyleLabel = document.createElement('span');
  validationErrorStyleLabel.textContent = t.validationErrorStyle;
  const validationErrorStyleSelect = createDialogSelect(
    [
      { value: 'stop', label: t.validationErrorStop },
      { value: 'warning', label: t.validationErrorWarning },
      { value: 'information', label: t.validationErrorInfo },
    ],
    'stop',
    { className: '', ariaLabel: t.validationErrorStyle },
  );
  validationErrorStyleRow.append(validationErrorStyleLabel, validationErrorStyleSelect);
  validationSection.appendChild(validationErrorStyleRow);

  const validationShowInputMessageRow = document.createElement('label');
  validationShowInputMessageRow.className = 'fc-fmtdlg__check';
  const validationShowInputMessageInput = document.createElement('input');
  validationShowInputMessageInput.type = 'checkbox';
  const validationShowInputMessageSpan = document.createElement('span');
  validationShowInputMessageSpan.textContent = t.validationShowInputMessage;
  validationShowInputMessageRow.append(
    validationShowInputMessageInput,
    validationShowInputMessageSpan,
  );
  validationSection.appendChild(validationShowInputMessageRow);

  const validationPromptTitleRow = document.createElement('label');
  validationPromptTitleRow.className = 'fc-fmtdlg__row';
  const validationPromptTitleLabel = document.createElement('span');
  validationPromptTitleLabel.textContent = t.validationPromptTitle;
  const validationPromptTitleInput = document.createElement('input');
  validationPromptTitleInput.type = 'text';
  validationPromptTitleInput.setAttribute('aria-label', t.validationPromptTitle);
  validationPromptTitleRow.append(validationPromptTitleLabel, validationPromptTitleInput);
  validationSection.appendChild(validationPromptTitleRow);

  const validationPromptMessageRow = document.createElement('div');
  validationPromptMessageRow.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
  const validationPromptMessageArea = document.createElement('textarea');
  validationPromptMessageArea.className = 'fc-fmtdlg__textarea';
  validationPromptMessageArea.rows = 3;
  validationPromptMessageArea.setAttribute('aria-label', t.validationPromptMessage);
  validationPromptMessageRow.appendChild(validationPromptMessageArea);
  validationSection.appendChild(validationPromptMessageRow);

  const validationShowErrorMessageRow = document.createElement('label');
  validationShowErrorMessageRow.className = 'fc-fmtdlg__check';
  const validationShowErrorMessageInput = document.createElement('input');
  validationShowErrorMessageInput.type = 'checkbox';
  const validationShowErrorMessageSpan = document.createElement('span');
  validationShowErrorMessageSpan.textContent = t.validationShowErrorMessage;
  validationShowErrorMessageRow.append(
    validationShowErrorMessageInput,
    validationShowErrorMessageSpan,
  );
  validationSection.appendChild(validationShowErrorMessageRow);

  const validationErrorTitleRow = document.createElement('label');
  validationErrorTitleRow.className = 'fc-fmtdlg__row';
  const validationErrorTitleLabel = document.createElement('span');
  validationErrorTitleLabel.textContent = t.validationErrorTitle;
  const validationErrorTitleInput = document.createElement('input');
  validationErrorTitleInput.type = 'text';
  validationErrorTitleInput.setAttribute('aria-label', t.validationErrorTitle);
  validationErrorTitleRow.append(validationErrorTitleLabel, validationErrorTitleInput);
  validationSection.appendChild(validationErrorTitleRow);

  const validationErrorMessageRow = document.createElement('div');
  validationErrorMessageRow.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
  const validationErrorMessageArea = document.createElement('textarea');
  validationErrorMessageArea.className = 'fc-fmtdlg__textarea';
  validationErrorMessageArea.rows = 3;
  validationErrorMessageArea.setAttribute('aria-label', t.validationErrorMessage);
  validationErrorMessageRow.appendChild(validationErrorMessageArea);
  validationSection.appendChild(validationErrorMessageRow);

  return {
    hyperlinkSection,
    commentSection,
    validationSection,
    hlInput,
    hlClear,
    commentArea,
    commentClear,
    validationKindSelect,
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
  };
}
