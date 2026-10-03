// Sheet tab: print area and titles (with reference validation), print
// options, comments / cell-error modes, and page order.
import { parsePrintAreas, parsePrintTitleCols, parsePrintTitleRows } from '../../commands/print.js';
import { colLetter, formatA1Range } from '../../engine/address.js';
import type {
  PageSetup,
  PrintCellErrorsMode,
  PrintCommentsMode,
  PrintPageOrder,
} from '../../store/store.js';
import { createDialogSelect } from '../../toolbar/dialogs/form-controls.js';
import { makeRow, makeTextInput } from '../page-setup-dialog-dom.js';
import { attachRangePickerButton } from '../range-picker-control.js';
import type { PageSetupTabContext } from './context.js';

export interface SheetTab {
  /** Validate the reference inputs; on failure show the error and focus the field. */
  validate(): boolean;
  hydrate(setup: PageSetup): void;
  collect(): Partial<PageSetup>;
}

export function createSheetTab(
  panel: HTMLDivElement,
  ctx: PageSetupTabContext,
  focusSheetTab: () => void,
): SheetTab {
  const { t, strings, store, on } = ctx;

  const makeOptionSelect = (
    label: string,
    options: { value: string; label: string }[],
  ): HTMLSelectElement => {
    return createDialogSelect(options, options[0]?.value ?? '', {
      className: 'fc-pgsetup__select',
      ariaLabel: label,
    });
  };

  // ── Sheet: print area / titles ──────────────────────────────────────────
  const printAreaRow = makeRow(t.printArea);
  const printAreaInput = makeTextInput('', t.printAreaPlaceholder);
  printAreaInput.setAttribute('aria-label', t.printArea);
  printAreaRow.valueCell.appendChild(printAreaInput);
  attachRangePickerButton(printAreaInput, {
    label: strings.pivotTableDialog.rangePickerSelect,
    getValue: () => formatA1Range(store.getState().selection.range),
    subscribeToRangeChanges: (listener) => store.subscribe(listener),
    kind: 'page-setup-print-area',
  });
  panel.appendChild(printAreaRow.row);

  // ── Print titles ────────────────────────────────────────────────────────
  const titleRowsRow = makeRow(t.printTitleRows);
  const titleRowsInput = makeTextInput('', t.printTitleRowsPlaceholder);
  titleRowsInput.setAttribute('aria-label', t.printTitleRows);
  titleRowsRow.valueCell.appendChild(titleRowsInput);
  attachRangePickerButton(titleRowsInput, {
    label: strings.pivotTableDialog.rangePickerSelect,
    getValue: () => {
      const range = store.getState().selection.range;
      return `${range.r0 + 1}:${range.r1 + 1}`;
    },
    subscribeToRangeChanges: (listener) => store.subscribe(listener),
    kind: 'page-setup-print-title-rows',
  });
  panel.appendChild(titleRowsRow.row);

  const titleColsRow = makeRow(t.printTitleCols);
  const titleColsInput = makeTextInput('', t.printTitleColsPlaceholder);
  titleColsInput.setAttribute('aria-label', t.printTitleCols);
  titleColsRow.valueCell.appendChild(titleColsInput);
  attachRangePickerButton(titleColsInput, {
    label: strings.pivotTableDialog.rangePickerSelect,
    getValue: () => {
      const range = store.getState().selection.range;
      return `${colLetter(range.c0)}:${colLetter(range.c1)}`;
    },
    subscribeToRangeChanges: (listener) => store.subscribe(listener),
    kind: 'page-setup-print-title-cols',
  });
  panel.appendChild(titleColsRow.row);

  // ── Sheet: print options ────────────────────────────────────────────────
  const gridRow = document.createElement('div');
  gridRow.className = 'fc-pgsetup__row fc-fmtdlg__row';
  const printOptionsTitle = document.createElement('span');
  printOptionsTitle.textContent = t.printOptions;
  const printOptionsValue = document.createElement('span');
  printOptionsValue.className = 'fc-pgsetup__value fc-pgsetup__checks';
  const showGridLabel = document.createElement('label');
  showGridLabel.className = 'fc-fmtdlg__check';
  const showGridInput = document.createElement('input');
  showGridInput.type = 'checkbox';
  const showGridText = document.createElement('span');
  showGridText.textContent = t.showGridlines;
  showGridLabel.append(showGridInput, showGridText);

  const showHeadLabel = document.createElement('label');
  showHeadLabel.className = 'fc-fmtdlg__check';
  const showHeadInput = document.createElement('input');
  showHeadInput.type = 'checkbox';
  const showHeadText = document.createElement('span');
  showHeadText.textContent = t.showHeadings;
  showHeadLabel.append(showHeadInput, showHeadText);

  const blackWhiteLabel = document.createElement('label');
  blackWhiteLabel.className = 'fc-fmtdlg__check';
  const blackWhiteInput = document.createElement('input');
  blackWhiteInput.type = 'checkbox';
  blackWhiteInput.setAttribute('aria-label', t.blackAndWhite);
  const blackWhiteText = document.createElement('span');
  blackWhiteText.textContent = t.blackAndWhite;
  blackWhiteLabel.append(blackWhiteInput, blackWhiteText);

  const draftLabel = document.createElement('label');
  draftLabel.className = 'fc-fmtdlg__check';
  const draftInput = document.createElement('input');
  draftInput.type = 'checkbox';
  draftInput.setAttribute('aria-label', t.draftQuality);
  const draftText = document.createElement('span');
  draftText.textContent = t.draftQuality;
  draftLabel.append(draftInput, draftText);

  printOptionsValue.append(showGridLabel, blackWhiteLabel, draftLabel, showHeadLabel);
  gridRow.append(printOptionsTitle, printOptionsValue);
  panel.appendChild(gridRow);

  const commentsRow = makeRow(t.comments);
  const commentsSelect = makeOptionSelect(t.comments, [
    { value: 'none', label: t.commentsNone },
    { value: 'asDisplayed', label: t.commentsAsDisplayed },
    { value: 'endOfSheet', label: t.commentsEndOfSheet },
  ]);
  commentsRow.valueCell.appendChild(commentsSelect);
  panel.appendChild(commentsRow.row);

  const errorsRow = makeRow(t.cellErrorsAs);
  const errorsSelect = makeOptionSelect(t.cellErrorsAs, [
    { value: 'displayed', label: t.cellErrorsDisplayed },
    { value: 'blank', label: t.cellErrorsBlank },
    { value: 'dash', label: t.cellErrorsDash },
    { value: 'na', label: t.cellErrorsNA },
  ]);
  errorsRow.valueCell.appendChild(errorsSelect);
  panel.appendChild(errorsRow.row);

  const pageOrderRow = document.createElement('div');
  pageOrderRow.className = 'fc-pgsetup__row fc-fmtdlg__row';
  const pageOrderTitle = document.createElement('span');
  pageOrderTitle.textContent = t.pageOrder;
  const pageOrderValue = document.createElement('span');
  pageOrderValue.className = 'fc-pgsetup__value';
  const downOverLabel = document.createElement('label');
  downOverLabel.className = 'fc-fmtdlg__check';
  const downOverInput = document.createElement('input');
  downOverInput.type = 'radio';
  downOverInput.name = 'fc-pgsetup-page-order';
  downOverInput.value = 'downThenOver';
  downOverInput.setAttribute('aria-label', t.pageOrderDownThenOver);
  const downOverText = document.createElement('span');
  downOverText.textContent = t.pageOrderDownThenOver;
  downOverLabel.append(downOverInput, downOverText);
  const overDownLabel = document.createElement('label');
  overDownLabel.className = 'fc-fmtdlg__check';
  const overDownInput = document.createElement('input');
  overDownInput.type = 'radio';
  overDownInput.name = 'fc-pgsetup-page-order';
  overDownInput.value = 'overThenDown';
  overDownInput.setAttribute('aria-label', t.pageOrderOverThenDown);
  const overDownText = document.createElement('span');
  overDownText.textContent = t.pageOrderOverThenDown;
  overDownLabel.append(overDownInput, overDownText);
  pageOrderValue.append(downOverLabel, overDownLabel);
  pageOrderRow.append(pageOrderTitle, pageOrderValue);
  panel.appendChild(pageOrderRow);

  const referenceError = document.createElement('div');
  referenceError.className = 'fc-pgsetup__error';
  referenceError.setAttribute('role', 'alert');
  referenceError.hidden = true;
  panel.appendChild(referenceError);

  const referenceInputs = [printAreaInput, titleRowsInput, titleColsInput] as const;

  const clearReferenceError = (): void => {
    referenceError.hidden = true;
    referenceError.textContent = '';
    for (const input of referenceInputs) input.removeAttribute('aria-invalid');
  };

  const showReferenceError = (input: HTMLInputElement, message: string): void => {
    for (const candidate of referenceInputs) candidate.removeAttribute('aria-invalid');
    input.setAttribute('aria-invalid', 'true');
    referenceError.textContent = message;
    referenceError.hidden = false;
    focusSheetTab();
    input.focus();
  };

  for (const input of referenceInputs) on(input, 'input', clearReferenceError);

  return {
    validate() {
      const area = printAreaInput.value.trim();
      if (area && !parsePrintAreas(area)) {
        showReferenceError(printAreaInput, t.invalidPrintArea);
        return false;
      }
      const titleRows = titleRowsInput.value.trim();
      if (titleRows && !parsePrintTitleRows(titleRows)) {
        showReferenceError(titleRowsInput, t.invalidPrintTitleRows);
        return false;
      }
      const titleCols = titleColsInput.value.trim();
      if (titleCols && !parsePrintTitleCols(titleCols)) {
        showReferenceError(titleColsInput, t.invalidPrintTitleCols);
        return false;
      }
      clearReferenceError();
      return true;
    },
    hydrate(setup) {
      printAreaInput.value = setup.printArea ?? '';
      titleRowsInput.value = setup.printTitleRows ?? '';
      titleColsInput.value = setup.printTitleCols ?? '';
      showGridInput.checked = setup.showGridlines === true;
      showHeadInput.checked = setup.showHeadings === true;
      blackWhiteInput.checked = setup.blackAndWhite === true;
      draftInput.checked = setup.draftQuality === true;
      commentsSelect.value = setup.comments ?? 'none';
      errorsSelect.value = setup.cellErrorsAs ?? 'displayed';
      downOverInput.checked = (setup.pageOrder ?? 'downThenOver') === 'downThenOver';
      overDownInput.checked = setup.pageOrder === 'overThenDown';
      clearReferenceError();
    },
    collect() {
      return {
        printArea: printAreaInput.value.trim() || undefined,
        printTitleRows: titleRowsInput.value.trim() || undefined,
        printTitleCols: titleColsInput.value.trim() || undefined,
        showGridlines: showGridInput.checked,
        showHeadings: showHeadInput.checked,
        blackAndWhite: blackWhiteInput.checked,
        draftQuality: draftInput.checked,
        comments: commentsSelect.value as PrintCommentsMode,
        cellErrorsAs: errorsSelect.value as PrintCellErrorsMode,
        pageOrder: (overDownInput.checked ? 'overThenDown' : 'downThenOver') as PrintPageOrder,
      };
    },
  };
}
